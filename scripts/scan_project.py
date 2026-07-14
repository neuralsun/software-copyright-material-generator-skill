#!/usr/bin/env python3
"""Scan a source tree and freeze a reproducible first-party source manifest.

The scanner is deliberately conservative: undecodable source is an error, secret
values are never copied into reports, and generated/vendor directories are ignored.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
from collections import Counter, defaultdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


DEFAULT_EXTENSIONS = {
    ".py", ".pyw", ".java", ".kt", ".kts", ".js", ".jsx", ".ts", ".tsx",
    ".vue", ".cs", ".go", ".rs", ".php", ".rb", ".swift", ".m", ".mm",
    ".c", ".h", ".cc", ".cpp", ".hpp", ".html", ".htm", ".css", ".scss",
    ".sass", ".less", ".sql", ".xml", ".xaml", ".dart", ".sh", ".ps1",
}

DEFAULT_EXCLUDED_PARTS = {
    ".git", ".svn", ".hg", ".idea", ".vscode", ".venv", "venv", "env",
    "node_modules", "vendor", "dist", "build", "target", "bin", "obj", ".next",
    ".nuxt", ".cache", "cache", "coverage", "htmlcov", "__pycache__", ".pytest_cache",
    ".mypy_cache", ".ruff_cache", ".runtime", ".runtime-logs", ".agents", "docs_gpt",
    "copyright-work", "previews", "uploads", "screenshots", "media", "materials", "材料",
}

EXTENSION_LANGUAGE = {
    ".py": "Python", ".pyw": "Python", ".java": "Java", ".kt": "Kotlin",
    ".kts": "Kotlin", ".js": "JavaScript", ".jsx": "JavaScript/JSX",
    ".ts": "TypeScript", ".tsx": "TypeScript/TSX", ".vue": "Vue SFC",
    ".cs": "C#", ".go": "Go", ".rs": "Rust", ".php": "PHP", ".rb": "Ruby",
    ".swift": "Swift", ".c": "C", ".h": "C/C++ header", ".cc": "C++",
    ".cpp": "C++", ".hpp": "C++ header", ".html": "HTML", ".htm": "HTML",
    ".css": "CSS", ".scss": "SCSS", ".sass": "Sass", ".less": "Less",
    ".sql": "SQL", ".xml": "XML", ".xaml": "XAML", ".dart": "Dart",
    ".sh": "Shell", ".ps1": "PowerShell",
}

LOGIN_HINTS = re.compile(r"(^|[/_.-])(auth|login|signin|signup|register)([/_.-]|$)", re.I)
MOJIBAKE_HINTS = re.compile(
    r"[�\ue000-\uf8ff]|(?:[鏈鐢璁鍖缃锛銆搴櫌涓鍦鐨鏄埛綍嶅瓨ㄣ].{0,3}){3,}"
)
PLACEHOLDER_HINTS = re.compile(r"\b(?:TODO|FIXME|XXX|HACK|TBD)\b|待补|待插入|后期补图|占位", re.I)
SECRET_ASSIGNMENT = re.compile(
    r"(?i)(api[_-]?key|access[_-]?key|secret(?:[_-]?key)?|client[_-]?secret|"
    r"password|passwd|pwd|token|invite[_-]?code|session[_-]?key)"
    r"[\"']?\s*(?:(?::\s*[A-Za-z_][\w.\[\]|, ]*\s*=)|[:=])\s*"
    r"[\"']([^\"']{6,})[\"']"
)
ENV_SECRET_DEFAULT = re.compile(
    r"(?i)(?:os\.)?(?:getenv|environ\.get)\(\s*[\"']"
    r"([^\"']*(?:key|secret|password|passwd|pwd|token|invite)[^\"']*)[\"']"
    r"\s*,\s*[\"']([^\"']{6,})[\"']"
)
DEFAULT_CREDENTIAL_LITERAL = re.compile(
    r"[\"']([A-Za-z][A-Za-z0-9_@.\-]*(?:12345678|123456|654321))[\"']",
    re.I,
)
PRIVATE_KEY_HINT = re.compile(r"-----BEGIN (?:RSA |EC |OPENSSH )?PRIVATE KEY-----")
ID_CARD_HINT = re.compile(r"(?<!\d)\d{17}[\dXx](?!\d)")
PHONE_HINT = re.compile(r"(?<!\d)1[3-9]\d{9}(?!\d)")
EMAIL_HINT = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")
VERSION_PATTERNS = [
    re.compile(r"(?i)\bapp_version\b(?:\s*:\s*[^=\r\n]+)?\s*=\s*[\"']([^\"']+)[\"']"),
    re.compile(r"(?i)[\"']?app_version[\"']?\s*:\s*[\"']([^\"']+)[\"']"),
    re.compile(r"(?i)\b(?:__version__|software_version)\b(?:\s*:\s*[^=\r\n]+)?\s*=\s*[\"']([^\"']+)[\"']"),
    re.compile(r"(?i)[\"']?(?:__version__|software_version)[\"']?\s*:\s*[\"']([^\"']+)[\"']"),
]
NAME_LITERAL = re.compile(r"[\"']([^\"'\r\n]{2,80}(?:系统|平台|软件))[^\"'\r\n]*[\"']")


def load_json(path: Path | None) -> dict[str, Any]:
    if path is None:
        return {}
    return json.loads(path.read_text(encoding="utf-8"))


def normalize_version(value: str) -> str:
    return value.strip().lstrip("vV")


def rel_posix(path: Path, root: Path) -> str:
    return path.relative_to(root).as_posix()


def decode_source(path: Path) -> tuple[str, str]:
    raw = path.read_bytes()
    if raw.startswith(b"\xef\xbb\xbf"):
        return raw.decode("utf-8-sig"), "utf-8-sig"
    if raw.startswith((b"\xff\xfe", b"\xfe\xff")):
        return raw.decode("utf-16"), "utf-16"
    for encoding in ("utf-8", "gb18030"):
        try:
            return raw.decode(encoding), encoding
        except UnicodeDecodeError:
            continue
    raise UnicodeDecodeError("source", raw, 0, min(len(raw), 1), "unsupported source encoding")


def is_excluded(path: Path, root: Path, config: dict[str, Any]) -> bool:
    relative = path.relative_to(root)
    lowered_parts = {part.lower() for part in relative.parts[:-1]}
    # Custom exclusions extend the safety defaults.  Replacing the defaults by
    # accident could silently admit node_modules/dist/vendor into a deposit.
    excluded = {part.lower() for part in DEFAULT_EXCLUDED_PARTS}
    excluded.update(str(x).lower() for x in config.get("exclude_parts", []))
    excluded.difference_update(str(x).lower() for x in config.get("allow_excluded_parts", []))
    if lowered_parts & excluded:
        return True
    name = path.name.lower()
    if name.endswith((".min.js", ".min.css", ".map", ".lock")):
        return True
    for pattern in config.get("exclude_globs", []):
        if relative.match(pattern):
            return True
    includes = config.get("include_globs", [])
    if includes and not any(relative.match(pattern) for pattern in includes):
        return True
    return False


def semantic_order_key(relative: str) -> tuple[int, str]:
    value = relative.lower()
    name = Path(value).name
    if name in {"main.py", "app.py", "program.cs", "main.go", "index.ts", "index.js"}:
        bucket = 10
    elif any(token in value for token in ("model", "entity", "schema", "database", "repository", "dao")):
        bucket = 20
    elif any(token in value for token in ("service", "domain", "core", "algorithm", "recommend", "workflow", "pipeline")):
        bucket = 30
    elif any(token in value for token in ("controller", "router", "route", "api", "handler")):
        bucket = 40
    elif LOGIN_HINTS.search(value) or any(token in value for token in ("security", "session", "permission")):
        bucket = 50
    elif any(token in value for token in ("config", "util", "common", "helper", "middleware")):
        bucket = 60
    elif any(token in value for token in ("view", "page", "screen", "component", "admin", "dashboard", "workspace")):
        bucket = 70
    elif path_suffix(relative) in {".css", ".scss", ".sass", ".less", ".html", ".htm"}:
        bucket = 80
    else:
        bucket = 65
    return bucket, value


def path_suffix(value: str) -> str:
    return Path(value).suffix.lower()


def read_package_facts(root: Path, excluded_config: dict[str, Any]) -> tuple[list[dict[str, Any]], dict[str, list[str]]]:
    versions: list[dict[str, Any]] = []
    dependencies: dict[str, list[str]] = defaultdict(list)
    for package_path in root.rglob("package.json"):
        if is_excluded(package_path, root, excluded_config):
            continue
        try:
            data = json.loads(package_path.read_text(encoding="utf-8"))
        except Exception:
            continue
        rel = rel_posix(package_path, root)
        if data.get("version"):
            versions.append({"value": str(data["version"]), "file": rel, "kind": "package.json"})
        for key in ("dependencies", "devDependencies"):
            dependencies["Node.js"].extend(sorted((data.get(key) or {}).keys()))
    for req in root.rglob("requirements*.txt"):
        if is_excluded(req, root, excluded_config):
            continue
        items: list[str] = []
        for line in req.read_text(encoding="utf-8-sig", errors="ignore").splitlines():
            clean = re.split(r"[<>=!~;\[]", line.strip(), maxsplit=1)[0]
            if clean and not clean.startswith(("#", "-")):
                items.append(clean)
        dependencies["Python"].extend(items)
    for pyproject in root.rglob("pyproject.toml"):
        if is_excluded(pyproject, root, excluded_config):
            continue
        try:
            import tomllib

            data = tomllib.loads(pyproject.read_text(encoding="utf-8"))
            project = data.get("project") or {}
            if project.get("version"):
                versions.append({"value": str(project["version"]), "file": rel_posix(pyproject, root), "kind": "pyproject"})
            for item in project.get("dependencies") or []:
                dependencies["Python"].append(re.split(r"[<>=!~;\[]", str(item), maxsplit=1)[0])
        except Exception:
            pass
    return versions, {k: sorted(set(v), key=str.lower) for k, v in dependencies.items()}


def extract_routes(relative: str, text: str) -> list[dict[str, Any]]:
    routes: list[dict[str, Any]] = []
    lines = text.splitlines()
    for index, line in enumerate(lines, start=1):
        py_match = re.search(r"@\w+\.(get|post|put|patch|delete)\(\s*[\"']([^\"']+)", line, re.I)
        if py_match:
            handler = ""
            for following in lines[index : min(index + 5, len(lines))]:
                def_match = re.search(r"(?:async\s+)?def\s+(\w+)", following)
                if def_match:
                    handler = def_match.group(1)
                    break
            routes.append({"kind": "backend", "method": py_match.group(1).upper(), "path": py_match.group(2), "handler": handler, "file": relative, "line": index})
        for js_match in re.finditer(r"(?:fetch|request)\(\s*[`\"']([^`\"']+)", line):
            routes.append({"kind": "frontend-api", "method": "", "path": js_match.group(1), "handler": "", "file": relative, "line": index})
        route_match = re.search(r"\bpath\s*:\s*[\"']([^\"']+)", line)
        if route_match:
            routes.append({"kind": "frontend-route", "method": "", "path": route_match.group(1), "handler": "", "file": relative, "line": index})
    return routes


def scan_risks(relative: str, text: str, expected_name: str, forbidden_terms: list[str]) -> list[dict[str, Any]]:
    risks: list[dict[str, Any]] = []
    for number, line in enumerate(text.splitlines(), start=1):
        if PLACEHOLDER_HINTS.search(line):
            risks.append({"severity": "warning", "type": "placeholder", "file": relative, "line": number})
        if MOJIBAKE_HINTS.search(line):
            risks.append({"severity": "error", "type": "mojibake", "file": relative, "line": number})
        if PRIVATE_KEY_HINT.search(line):
            risks.append({"severity": "error", "type": "private-key", "file": relative, "line": number})
        secret_reported = False
        for match in SECRET_ASSIGNMENT.finditer(line):
            value = match.group(2).strip()
            if value.lower() not in {"changeme", "change-me", "password", "your-key", "example", "placeholder"}:
                risks.append({"severity": "error", "type": "hardcoded-secret-or-credential", "name": match.group(1), "file": relative, "line": number, "value": "<redacted>"})
                secret_reported = True
        for match in ENV_SECRET_DEFAULT.finditer(line):
            value = match.group(2).strip()
            if value.lower() not in {"changeme", "change-me", "password", "your-key", "example", "placeholder"}:
                risks.append({"severity": "error", "type": "hardcoded-secret-or-credential", "name": match.group(1), "file": relative, "line": number, "value": "<redacted>"})
                secret_reported = True
        if not secret_reported and DEFAULT_CREDENTIAL_LITERAL.search(line):
            risks.append({"severity": "error", "type": "password-like-literal", "name": "password-like-literal", "file": relative, "line": number, "value": "<redacted>"})
        if ID_CARD_HINT.search(line):
            risks.append({"severity": "warning", "type": "possible-id-card", "file": relative, "line": number})
        if PHONE_HINT.search(line):
            risks.append({"severity": "warning", "type": "possible-mobile", "file": relative, "line": number})
        if EMAIL_HINT.search(line) and not re.search(r"example\.(?:com|org|net)", line, re.I):
            risks.append({"severity": "warning", "type": "possible-email", "file": relative, "line": number})
        for term in forbidden_terms:
            if term and term in line:
                risks.append({"severity": "error", "type": "forbidden-term", "term": term, "file": relative, "line": number})
        if expected_name and ("系统" in line or "平台" in line) and expected_name not in line:
            # Name candidates are reported separately; do not mark every UI phrase as an error.
            pass
    return risks


def main() -> int:
    parser = argparse.ArgumentParser(description="Scan first-party source and write project facts plus a reproducible source manifest.")
    parser.add_argument("--project", required=True, type=Path, help="Project root")
    parser.add_argument("--output", required=True, type=Path, help="project_facts.json output")
    parser.add_argument("--manifest", type=Path, help="source_manifest.json output (defaults next to output)")
    parser.add_argument("--config", type=Path, help="Optional JSON configuration")
    parser.add_argument("--expected-name", default="", help="Registered full software name")
    parser.add_argument("--expected-version", default="", help="Registered version, e.g. V1.0")
    args = parser.parse_args()

    root = args.project.resolve()
    if not root.is_dir():
        parser.error(f"Project root does not exist: {root}")
    config = load_json(args.config)
    source_config = config.get("source") or config
    expected_name = args.expected_name or ((config.get("software") or {}).get("full_name") or "")
    expected_version = args.expected_version or ((config.get("software") or {}).get("version") or "")
    extensions = {str(x).lower() for x in source_config.get("include_extensions", DEFAULT_EXTENSIONS)}
    forbidden_terms = [str(x) for x in source_config.get("forbidden_terms", [])]

    raw_explicit_order = [str(x).replace("\\", "/") for x in source_config.get("source_order", [])]
    explicit_order = list(dict.fromkeys(raw_explicit_order))
    discovery_errors: list[dict[str, Any]] = []
    if len(explicit_order) != len(raw_explicit_order):
        duplicates = sorted({item for item in raw_explicit_order if raw_explicit_order.count(item) > 1})
        discovery_errors.append({"type": "duplicate-explicit-source", "files": duplicates})
    discovered: list[Path] = []
    for path in root.rglob("*"):
        if not path.is_file() or path.suffix.lower() not in extensions:
            continue
        if is_excluded(path, root, source_config):
            continue
        if path.is_symlink():
            discovery_errors.append({"type": "symlink-source", "file": rel_posix(path, root)})
            continue
        discovered.append(path)

    by_rel = {rel_posix(path, root): path for path in discovered}
    missing_explicit = [item for item in explicit_order if item not in by_rel]
    ordered_rels = [item for item in explicit_order if item in by_rel]
    remaining = sorted((item for item in by_rel if item not in set(ordered_rels)), key=semantic_order_key)
    ordered_rels.extend(remaining)

    files: list[dict[str, Any]] = []
    errors: list[dict[str, Any]] = list(discovery_errors)
    risks: list[dict[str, Any]] = []
    versions, dependencies = read_package_facts(root, source_config)
    name_candidates: Counter[str] = Counter()
    routes: list[dict[str, Any]] = []
    language_lines: Counter[str] = Counter()
    program_lines = 0
    snapshot = hashlib.sha256()

    for relative in ordered_rels:
        path = by_rel[relative]
        try:
            text, encoding = decode_source(path)
        except UnicodeDecodeError as error:
            errors.append({"type": "decode-error", "file": relative, "message": str(error)})
            continue
        if "�" in text:
            errors.append({"type": "replacement-character", "file": relative})
            continue
        physical_lines = text.splitlines()
        nonblank = sum(1 for line in physical_lines if line.strip())
        digest = hashlib.sha256(path.read_bytes()).hexdigest()
        language = EXTENSION_LANGUAGE.get(path.suffix.lower(), path.suffix.lower().lstrip(".").upper())
        login_ui = bool(LOGIN_HINTS.search(relative)) and path.suffix.lower() in {".html", ".htm", ".vue", ".jsx", ".tsx", ".css", ".scss", ".less"}
        files.append({
            "path": relative,
            "language": language,
            "encoding": encoding,
            "physical_lines": len(physical_lines),
            "nonblank_lines": nonblank,
            "sha256": digest,
            "login_ui_risk": login_ui,
        })
        program_lines += nonblank
        language_lines[language] += nonblank
        snapshot.update(f"{relative}\0{digest}\0{nonblank}\n".encode("utf-8"))
        routes.extend(extract_routes(relative, text))
        risks.extend(scan_risks(relative, text, expected_name, forbidden_terms))
        for pattern in VERSION_PATTERNS:
            for match in pattern.finditer(text):
                versions.append({"value": match.group(1), "file": relative, "line": text.count("\n", 0, match.start()) + 1, "kind": "source"})
        for candidate in NAME_LITERAL.findall(text):
            cleaned = re.sub(r"\s+", "", candidate).strip("：:，,。.;；")
            if 3 <= len(cleaned) <= 60:
                name_candidates[cleaned] += 1

    if missing_explicit:
        errors.append({"type": "missing-explicit-source", "files": missing_explicit})

    if not files or program_lines <= 0:
        errors.append({"type": "empty-source-set", "message": "No nonblank first-party source lines were discovered"})

    # A redacted risk record is safe to persist, but high-severity findings must
    # also affect the final readiness flag.  Previously private keys and literal
    # credentials were reported while submission_ready incorrectly stayed true.
    for risk in risks:
        if risk.get("severity") == "error":
            errors.append({
                "type": "blocking-source-risk",
                "risk_type": risk.get("type"),
                "file": risk.get("file"),
                "line": risk.get("line"),
            })

    unique_versions = sorted({normalize_version(item["value"]) for item in versions if str(item.get("value", "")).strip()})
    expected_normalized = normalize_version(expected_version) if expected_version else ""
    version_mismatch = bool(expected_normalized and any(value != expected_normalized for value in unique_versions))
    if version_mismatch:
        errors.append({"type": "version-mismatch", "expected": expected_normalized, "found": unique_versions})

    manifest = {
        "schema_version": 1,
        "project_root": str(root),
        "created_at": datetime.now(timezone.utc).isoformat(),
        "snapshot_sha256": snapshot.hexdigest(),
        "program_nonblank_lines": program_lines,
        "source_policy": {
            "include_extensions": sorted(extensions),
            "excluded_parts": sorted(
                ({part.lower() for part in DEFAULT_EXCLUDED_PARTS}
                 | {str(x).lower() for x in source_config.get("exclude_parts", [])})
                - {str(x).lower() for x in source_config.get("allow_excluded_parts", [])}
            ),
            "exclude_globs": [str(x) for x in source_config.get("exclude_globs", [])],
            "include_globs": [str(x) for x in source_config.get("include_globs", [])],
            "excluded_suffixes": [".min.js", ".min.css", ".map", ".lock"],
        },
        "files": files,
    }
    facts = {
        "schema_version": 1,
        "software": {"expected_name": expected_name, "expected_version": expected_version},
        "snapshot_sha256": manifest["snapshot_sha256"],
        "program_nonblank_lines": program_lines,
        "source_file_count": len(files),
        "language_lines": dict(language_lines.most_common()),
        "versions": versions,
        "version_values": unique_versions,
        "name_candidates": [{"value": value, "occurrences": count} for value, count in name_candidates.most_common(30)],
        "dependencies": dependencies,
        "routes": routes,
        "login_ui_files": [item["path"] for item in files if item["login_ui_risk"]],
        "risks": risks,
        "errors": errors,
        "submission_ready": not errors,
        "capability_boundary": "Static evidence only. Run isolated build/smoke tests before describing a feature as complete.",
    }

    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(facts, ensure_ascii=False, indent=2), encoding="utf-8")
    manifest_path = args.manifest or args.output.with_name("source_manifest.json")
    manifest_path.parent.mkdir(parents=True, exist_ok=True)
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"facts": str(args.output), "manifest": str(manifest_path), "files": len(files), "program_lines": program_lines, "errors": len(errors), "warnings": sum(1 for item in risks if item["severity"] == "warning")}, ensure_ascii=False))
    return 0 if not errors else 2


if __name__ == "__main__":
    sys.exit(main())
