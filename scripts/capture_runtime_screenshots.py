#!/usr/bin/env python3
"""Capture genuine local runtime screens with Playwright Chromium.

Playwright is deliberately optional.  When it is unavailable, this script
validates the scenario file and writes a manual screenshot checklist instead
of inventing UI evidence.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import sys
from datetime import datetime
from pathlib import Path
from typing import Any
from urllib.parse import urljoin, urlparse


SUPPORTED_ACTIONS = {
    "goto", "fill", "click", "select", "check", "upload",
    "wait_for", "wait_for_url", "screenshot",
}
REVIEW_FIELDS = (
    "software_name_correct",
    "software_version_correct",
    "old_name_absent",
    "pii_absent_or_redacted",
    "credential_absent",
    "secret_absent",
    "error_absent",
    "splice_absent",
    "feature_evidence_supported",
    "manually_reviewed",
)


class ScenarioError(ValueError):
    """Raised for an invalid screenshot scenario."""


def load_json(path: Path) -> dict[str, Any]:
    value = json.loads(path.read_text(encoding="utf-8-sig"))
    if not isinstance(value, dict):
        raise ScenarioError("scenario JSON root must be an object")
    return value


def validate_scenarios(payload: dict[str, Any]) -> list[dict[str, Any]]:
    scenarios = payload.get("scenarios")
    if not isinstance(scenarios, list) or not scenarios:
        raise ScenarioError("scenarios must be a non-empty array")
    seen: set[str] = set()
    result: list[dict[str, Any]] = []
    for index, raw in enumerate(scenarios, start=1):
        if not isinstance(raw, dict):
            raise ScenarioError(f"scenario #{index} must be an object")
        scenario_id = str(raw.get("id") or "").strip()
        if not scenario_id or scenario_id in seen:
            raise ScenarioError(f"scenario #{index} has a missing or duplicate id")
        seen.add(scenario_id)
        features = raw.get("feature_ids")
        if not isinstance(features, list) or not any(str(v).strip() for v in features):
            raise ScenarioError(f"scenario {scenario_id} requires feature_ids")
        steps = raw.get("steps")
        if not isinstance(steps, list) or not steps:
            raise ScenarioError(f"scenario {scenario_id} requires steps")
        if not any(isinstance(s, dict) and s.get("action") == "screenshot" for s in steps):
            raise ScenarioError(f"scenario {scenario_id} has no screenshot step")
        for position, step in enumerate(steps, start=1):
            if not isinstance(step, dict):
                raise ScenarioError(f"scenario {scenario_id} step {position} must be an object")
            action = str(step.get("action") or "").strip()
            if action not in SUPPORTED_ACTIONS:
                raise ScenarioError(f"scenario {scenario_id} step {position}: unsupported action {action!r}")
            if action in {"fill", "click", "select", "check", "upload", "wait_for"} and not step.get("locator"):
                raise ScenarioError(f"scenario {scenario_id} step {position}: {action} requires locator")
            if action in {"goto", "wait_for_url", "fill", "select", "upload"} and step.get("value") is None:
                raise ScenarioError(f"scenario {scenario_id} step {position}: {action} requires value")
            if action == "screenshot" and not str(step.get("filename") or "").strip():
                raise ScenarioError(f"scenario {scenario_id} step {position}: screenshot filename is required")
            if action == "screenshot" and not Path(str(step.get("filename"))).name.startswith(f"{scenario_id}_"):
                raise ScenarioError(f"scenario {scenario_id} step {position}: filename must start with {scenario_id}_")
        result.append(raw)
    return result


def bool_arg(value: str) -> bool:
    lowered = value.strip().lower()
    if lowered in {"1", "true", "yes", "on"}:
        return True
    if lowered in {"0", "false", "no", "off"}:
        return False
    raise argparse.ArgumentTypeError("expected true or false")


def secret_value(value: Any) -> str:
    if isinstance(value, dict) and isinstance(value.get("env"), str):
        name = value["env"]
        if name not in os.environ:
            raise ScenarioError(f"required environment variable is not set: {name}")
        return os.environ[name]
    text = str(value if value is not None else "")
    match = re.fullmatch(r"(?:\$ENV\{|env:)([A-Za-z_][A-Za-z0-9_]*)\}?", text)
    if match:
        name = match.group(1)
        if name not in os.environ:
            raise ScenarioError(f"required environment variable is not set: {name}")
        return os.environ[name]
    return text


def resolve_url(base_url: str, value: Any) -> str:
    text = str(value or "").strip()
    if not text:
        raise ScenarioError("navigation URL/path cannot be empty")
    resolved = urljoin(base_url.rstrip("/") + "/", text.lstrip("/"))
    parsed = urlparse(resolved)
    if parsed.scheme not in {"http", "https"} or parsed.hostname not in {"127.0.0.1", "localhost", "::1"}:
        raise ScenarioError(f"screenshot automation only permits local isolated runtime URLs: {resolved}")
    return resolved


def make_locator(page: Any, spec: Any) -> Any:
    if isinstance(spec, str):
        return page.get_by_text(spec)
    if not isinstance(spec, dict):
        raise ScenarioError("locator must be a string or object")
    kind = str(spec.get("type") or "").strip().lower()
    value = str(spec.get("value") or "").strip()
    exact = bool(spec.get("exact", False))
    if not kind or not value:
        raise ScenarioError("locator.type and locator.value are required")
    if kind == "role":
        name = spec.get("name")
        return page.get_by_role(value, name=name, exact=exact)
    if kind == "label":
        return page.get_by_label(value, exact=exact)
    if kind == "placeholder":
        return page.get_by_placeholder(value, exact=exact)
    if kind == "text":
        return page.get_by_text(value, exact=exact)
    if kind in {"test_id", "testid"}:
        return page.get_by_test_id(value)
    if kind in {"css", "xpath", "selector"}:
        selector = f"xpath={value}" if kind == "xpath" else value
        return page.locator(selector)
    raise ScenarioError(f"unsupported locator type: {kind}")


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def review_template() -> dict[str, bool]:
    # Automation must never approve human-review fields.
    return {field: False for field in REVIEW_FIELDS}


def installed_chromium_executable(explicit: Path | None) -> Path | None:
    candidates: list[Path] = []
    if explicit:
        candidates.append(explicit.expanduser())
    env_path = os.environ.get("PLAYWRIGHT_CHROMIUM_EXECUTABLE")
    if env_path:
        candidates.append(Path(env_path).expanduser())
    for root_name in ("PROGRAMFILES", "PROGRAMFILES(X86)", "LOCALAPPDATA"):
        root = os.environ.get(root_name)
        if not root:
            continue
        candidates.extend([
            Path(root) / "Google/Chrome/Application/chrome.exe",
            Path(root) / "Microsoft/Edge/Application/msedge.exe",
        ])
    for candidate in candidates:
        resolved = candidate.resolve()
        if resolved.is_file():
            return resolved
    return None


def write_manifest(
    path: Path,
    entries: list[dict[str, Any]],
    failures: list[dict[str, Any]],
    existing: dict[str, Any] | None = None,
) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    previous = existing.get("figures", []) if isinstance(existing, dict) else []
    replacement_ids = {str(item.get("id")) for item in entries}
    retained = [
        item for item in previous
        if isinstance(item, dict) and str(item.get("id")) not in replacement_ids
    ] if isinstance(previous, list) else []
    payload = {
        **(existing or {}),
        "schema_version": 2,
        "generated_at": datetime.now().astimezone().isoformat(timespec="seconds"),
        "review_notice": "自动采集不等于人工复核；逐项复核后方可手工修改 review 字段。",
        "figures": retained + entries,
        "capture_failures": failures,
    }
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def write_manual_checklist(
    checklist: Path,
    scenarios: list[dict[str, Any]],
    width: int,
    height: int,
    reason: str,
) -> None:
    lines = [
        "# 人工截图清单（材料保持草稿状态）", "",
        f"> 自动截图不可用原因：{reason}",
        "> 该原因是采集环境/工具状态，不代表被测软件存在功能错误。", "",
        f"建议浏览器视口：{width}×{height}。截图必须来自 localhost/127.0.0.1 隔离测试环境。", "",
    ]
    for item in scenarios:
        shots = [s for s in item["steps"] if s.get("action") == "screenshot"]
        route = item.get("operation_path") or item.get("path") or "按场景步骤进入"
        test_data = item.get("test_data") or "准备虚构、脱敏测试数据"
        filenames = "、".join(str(s.get("filename")) for s in shots)
        lines.extend([
            f"## {item['id']} {item.get('name', '')}", "",
            f"- 功能 ID：{', '.join(str(v) for v in item['feature_ids'])}",
            f"- 操作路径：{route}",
            f"- 建议文件名：{filenames}",
            f"- 测试数据：{test_data}",
            "- 隐私处理：隐藏真实姓名、手机号、证件号、地址、账号密码、API Key、Token、数据库连接与公网地址。",
            "- 复核：软件全称/版本正确；无旧名称、隐私、凭据、密钥、报错、404/500、调试信息、拼接涂改；画面能够证明所映射功能。", "",
        ])
    lines.extend([
        "## 完成条件", "",
        "将截图放入 `copyright-work/figures/`，补录 `figures_manifest.json` 的来源、URL、视口、操作路径、测试数据和 SHA-256。",
        "所有 review 项只能在人工确认后手工设为 true；全部正式截图通过后，才可在主配置中手工设置 `screenshots_reviewed=true`。", "",
    ])
    checklist.parent.mkdir(parents=True, exist_ok=True)
    checklist.write_text("\n".join(lines), encoding="utf-8")


def run_step(page: Any, step: dict[str, Any], base_url: str, timeout: int) -> None:
    action = step["action"]
    if action == "goto":
        page.goto(resolve_url(base_url, step.get("value")), wait_until=step.get("wait_until", "domcontentloaded"), timeout=timeout)
    elif action == "wait_for_url":
        page.wait_for_url(resolve_url(base_url, step.get("value")), timeout=timeout)
    elif action == "screenshot":
        return
    else:
        locator = make_locator(page, step.get("locator"))
        if action == "wait_for":
            locator.wait_for(state=step.get("state", "visible"), timeout=timeout)
        elif action == "fill":
            locator.fill(secret_value(step.get("value")), timeout=timeout)
        elif action == "click":
            locator.click(timeout=timeout)
        elif action == "select":
            locator.select_option(secret_value(step.get("value")), timeout=timeout)
        elif action == "check":
            locator.check(timeout=timeout)
        elif action == "upload":
            upload = Path(secret_value(step.get("value"))).expanduser().resolve()
            if not upload.is_file():
                raise ScenarioError(f"upload file does not exist: {upload}")
            locator.set_input_files(str(upload), timeout=timeout)


def capture(args: argparse.Namespace, payload: dict[str, Any], scenarios: list[dict[str, Any]]) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    try:
        from playwright.sync_api import sync_playwright
    except ImportError as exc:
        raise RuntimeError(
            "Python Playwright 未安装。请执行：python -m pip install -r requirements-screenshot.txt；"
            "然后执行：python -m playwright install chromium"
        ) from exc

    base_url = args.base_url or str(payload.get("base_url") or "").strip()
    if not base_url:
        raise ScenarioError("--base-url or scenario base_url is required")
    resolve_url(base_url, "/")
    entries: list[dict[str, Any]] = []
    failures: list[dict[str, Any]] = []
    args.output_dir.mkdir(parents=True, exist_ok=True)
    with sync_playwright() as playwright:
        executable = installed_chromium_executable(args.executable_path)
        launch_options: dict[str, Any] = {"headless": args.headless}
        if executable:
            launch_options["executable_path"] = str(executable)
        try:
            browser = playwright.chromium.launch(**launch_options)
        except Exception as exc:
            raise RuntimeError(
                "Playwright 无法启动 Chromium。请执行：python -m playwright install chromium；"
                "或用 --executable-path 指向已安装的 Chromium/Chrome/Edge。"
            ) from exc
        context_options: dict[str, Any] = {"viewport": {"width": args.viewport_width, "height": args.viewport_height}}
        if args.storage_state:
            context_options["storage_state"] = str(args.storage_state.expanduser().resolve())
        context = browser.new_context(**context_options)
        try:
            for scenario in scenarios:
                page = context.new_page()
                try:
                    if scenario.get("path"):
                        page.goto(resolve_url(base_url, scenario["path"]), wait_until="domcontentloaded", timeout=args.timeout)
                    shot_number = 0
                    for step_number, step in enumerate(scenario["steps"], start=1):
                        try:
                            if step["action"] != "screenshot":
                                run_step(page, step, base_url, args.timeout)
                                continue
                            shot_number += 1
                            target = make_locator(page, step["locator"]) if step.get("locator") else None
                            if target is not None:
                                target.wait_for(state="visible", timeout=args.timeout)
                            filename = Path(str(step["filename"])).name
                            output = (args.output_dir / filename).resolve()
                            if output.suffix.lower() not in {".png", ".jpg", ".jpeg"}:
                                raise ScenarioError("screenshot filename must end in .png, .jpg or .jpeg")
                            # Viewport evidence is the default; locator is only an assertion unless capture_element=true.
                            if target is not None and bool(step.get("capture_element", False)):
                                target.screenshot(path=str(output), timeout=args.timeout)
                            else:
                                page.screenshot(path=str(output), full_page=False, timeout=args.timeout)
                            figure_id = str(step.get("id") or (scenario["id"] if shot_number == 1 else f"{scenario['id']}_{shot_number}"))
                            try:
                                relative = output.relative_to(args.manifest.parent.resolve()).as_posix()
                            except ValueError:
                                relative = str(output)
                            entries.append({
                                "id": figure_id,
                                "path": relative,
                                "kind": "runtime_screenshot",
                                "feature_ids": [str(v) for v in scenario["feature_ids"]],
                                "caption": str(step.get("caption") or scenario.get("name") or figure_id),
                                "captured_at": datetime.now().astimezone().isoformat(timespec="seconds"),
                                "source": "local isolated runtime",
                                "capture_tool": "playwright_chromium",
                                "runtime_url": page.url,
                                "viewport": f"{args.viewport_width}x{args.viewport_height}",
                                "scenario_id": str(scenario["id"]),
                                "operation_path": str(scenario.get("operation_path") or scenario.get("path") or "按场景步骤执行"),
                                "test_data": str(scenario.get("test_data") or "虚构、脱敏测试数据；具体值不写入清单"),
                                "sha256": sha256_file(output),
                                "review": review_template(),
                            })
                        except Exception as exc:  # keep other scenarios/screens
                            failures.append({"scenario_id": scenario["id"], "step": step_number, "action": step.get("action"), "error": str(exc)})
                            break
                except Exception as exc:
                    failures.append({"scenario_id": scenario["id"], "step": 0, "action": "open", "error": str(exc)})
                finally:
                    page.close()
        finally:
            context.close()
            browser.close()
    return entries, failures


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Capture genuine local runtime screenshots with Playwright Chromium.")
    parser.add_argument("--base-url", help="local runtime URL; overrides scenario JSON")
    parser.add_argument("--output-dir", required=True, type=Path)
    parser.add_argument("--manifest", required=True, type=Path)
    parser.add_argument("--scenario", required=True, type=Path, help="screenshot_scenarios.json")
    parser.add_argument("--headless", type=bool_arg, default=True, metavar="true|false")
    parser.add_argument("--viewport-width", type=int, default=1440)
    parser.add_argument("--viewport-height", type=int, default=900)
    parser.add_argument("--timeout", type=int, default=30000, help="milliseconds")
    parser.add_argument("--storage-state", type=Path)
    parser.add_argument("--executable-path", type=Path, help="optional installed Chromium/Chrome/Edge executable; auto-detected on Windows")
    parser.add_argument("--check-only", action="store_true", help="validate scenarios without launching a browser")
    parser.add_argument("--manual-checklist", type=Path, help="fallback checklist path (default beside manifest)")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    if args.viewport_width < 320 or args.viewport_height < 240 or args.timeout <= 0:
        raise ScenarioError("viewport is too small or timeout is not positive")
    args.scenario = args.scenario.expanduser().resolve()
    args.output_dir = args.output_dir.expanduser().resolve()
    args.manifest = args.manifest.expanduser().resolve()
    if args.storage_state:
        args.storage_state = args.storage_state.expanduser().resolve()
        if not args.storage_state.is_file():
            raise ScenarioError(f"--storage-state file does not exist: {args.storage_state}")
    payload = load_json(args.scenario)
    scenarios = validate_scenarios(payload)
    if args.check_only:
        print(json.dumps({"valid": True, "scenarios": len(scenarios)}, ensure_ascii=False))
        return 0
    checklist = (args.manual_checklist or (args.manifest.parent / "manual_screenshot_checklist.md")).expanduser().resolve()
    existing_manifest: dict[str, Any] | None = None
    if args.manifest.exists():
        existing_manifest = load_json(args.manifest)
    try:
        entries, failures = capture(args, payload, scenarios)
    except RuntimeError as exc:
        write_manual_checklist(checklist, scenarios, args.viewport_width, args.viewport_height, str(exc))
        print(f"error: {exc}\n人工截图清单：{checklist}", file=sys.stderr)
        return 3
    write_manifest(args.manifest, entries, failures, existing_manifest)
    if failures:
        write_manual_checklist(checklist, scenarios, args.viewport_width, args.viewport_height, "部分自动截图步骤失败；详见 figures_manifest.json 的 capture_failures")
    print(json.dumps({"screenshots": len(entries), "failures": len(failures), "manifest": str(args.manifest)}, ensure_ascii=False))
    return 2 if failures else 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (ScenarioError, FileNotFoundError, json.JSONDecodeError) as exc:
        print(f"error: {exc}", file=sys.stderr)
        raise SystemExit(2)
