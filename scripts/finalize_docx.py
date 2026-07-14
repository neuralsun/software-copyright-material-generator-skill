#!/usr/bin/env python3
"""Finalize DOCX package metadata and cached PAGE/NUMPAGES results.

This script does not calculate layout. First repaginate with Word/WPS/LibreOffice,
then pass the measured page count here. It intentionally never guesses page count.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import tempfile
from datetime import datetime, timezone
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
DC = "http://purl.org/dc/elements/1.1/"
CP = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
DCTERMS = "http://purl.org/dc/terms/"
EP = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
XSI = "http://www.w3.org/2001/XMLSchema-instance"
XML_PARSER = etree.XMLParser(resolve_entities=False, no_network=True, huge_tree=False)


def parse_xml(data: bytes) -> etree._Element:
    return etree.fromstring(data, parser=XML_PARSER)


def is_field(instruction: str, field_name: str) -> bool:
    return bool(re.search(rf"(?<![A-Z]){re.escape(field_name.upper())}(?![A-Z])", instruction.upper()))


def set_field_result(root: etree._Element, field_name: str, value: str) -> int:
    """Update cached results for complex and simple Word fields.

    Word may split an instruction across several ``w:instrText`` nodes.  A
    small field-state stack avoids changing a neighbouring PAGE field when a
    paragraph contains both PAGE and NUMPAGES.
    """
    changed = 0
    for paragraph in root.iter(f"{{{W}}}p"):
        stack: list[dict[str, object]] = []
        for node in paragraph.iter():
            if node.tag == f"{{{W}}}fldChar":
                kind = node.get(f"{{{W}}}fldCharType")
                if kind == "begin":
                    stack.append({"instruction": [], "separated": False, "match": False, "updated": False})
                elif kind == "separate" and stack:
                    context = stack[-1]
                    context["separated"] = True
                    context["match"] = is_field("".join(context["instruction"]), field_name)  # type: ignore[arg-type]
                elif kind == "end" and stack:
                    stack.pop()
                continue
            if node.tag == f"{{{W}}}instrText" and stack and not stack[-1]["separated"]:
                stack[-1]["instruction"].append(node.text or "")  # type: ignore[union-attr]
                continue
            if node.tag == f"{{{W}}}t":
                for context in reversed(stack):
                    if context["separated"] and context["match"] and not context["updated"]:
                        node.text = value
                        context["updated"] = True
                        changed += 1
                        break

    for simple in root.iter(f"{{{W}}}fldSimple"):
        if not is_field(simple.get(f"{{{W}}}instr", ""), field_name):
            continue
        text_node = next(simple.iter(f"{{{W}}}t"), None)
        if text_node is None:
            run = etree.SubElement(simple, f"{{{W}}}r")
            text_node = etree.SubElement(run, f"{{{W}}}t")
        text_node.text = value
        changed += 1
    return changed


def set_update_fields(root: etree._Element) -> None:
    node = root.find(f"{{{W}}}updateFields")
    if node is None:
        node = etree.SubElement(root, f"{{{W}}}updateFields")
    node.set(f"{{{W}}}val", "true")


def set_or_create(parent: etree._Element, tag: str, value: str) -> None:
    node = parent.find(tag)
    if node is None:
        node = etree.SubElement(parent, tag)
    node.text = value


def finalize(path: Path, pages: int, title: str = "") -> dict[str, object]:
    path = path.resolve()
    if pages <= 0:
        raise ValueError("Measured page count must be positive")
    timestamp = datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")
    fd, temp_name = tempfile.mkstemp(prefix=f"{path.stem}_", suffix=".docx", dir=path.parent)
    os.close(fd)
    temp = Path(temp_name)
    field_updates = 0
    try:
        with ZipFile(path, "r") as source, ZipFile(temp, "w", compression=ZIP_DEFLATED) as target:
            if source.testzip() is not None:
                raise ValueError(f"Corrupt DOCX package: {path}")
            for info in source.infolist():
                data = source.read(info.filename)
                if info.filename == "docProps/core.xml":
                    root = parse_xml(data)
                    set_or_create(root, f"{{{DC}}}creator", "")
                    set_or_create(root, f"{{{CP}}}lastModifiedBy", "")
                    # Never retain a title or descriptive metadata inherited
                    # from an older project/template.  A supplied title wins;
                    # otherwise these optional fields are deliberately blank.
                    set_or_create(root, f"{{{DC}}}title", title)
                    for tag in (
                        f"{{{DC}}}subject",
                        f"{{{DC}}}description",
                        f"{{{DC}}}identifier",
                        f"{{{CP}}}keywords",
                        f"{{{CP}}}category",
                        f"{{{CP}}}contentStatus",
                    ):
                        set_or_create(root, tag, "")
                    for tag in (f"{{{DCTERMS}}}created", f"{{{DCTERMS}}}modified"):
                        set_or_create(root, tag, timestamp)
                        root.find(tag).set(f"{{{XSI}}}type", "dcterms:W3CDTF")
                    data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
                elif info.filename == "docProps/app.xml":
                    root = parse_xml(data)
                    set_or_create(root, f"{{{EP}}}Pages", str(pages))
                    for tag in (f"{{{EP}}}Company", f"{{{EP}}}Manager", f"{{{EP}}}HyperlinkBase"):
                        set_or_create(root, tag, "")
                    data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
                elif info.filename == "word/settings.xml":
                    root = parse_xml(data)
                    set_update_fields(root)
                    data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
                elif info.filename.endswith(".xml") and (
                    info.filename == "word/document.xml"
                    or info.filename.startswith("word/header")
                    or info.filename.startswith("word/footer")
                    or info.filename in {"word/footnotes.xml", "word/endnotes.xml"}
                ):
                    root = parse_xml(data)
                    field_updates += set_field_result(root, "NUMPAGES", str(pages))
                    data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
                target.writestr(info, data)
        os.replace(temp, path)
    finally:
        if temp.exists():
            temp.unlink()
    return {"path": str(path), "pages": pages, "numpages_cache_updates": field_updates}


def main() -> int:
    parser = argparse.ArgumentParser(description="Patch measured page count, field caches, updateFields and clean DOCX metadata.")
    parser.add_argument("--path", action="append", required=True, type=Path, help="DOCX path; repeat for multiple files")
    parser.add_argument("--pages", action="append", required=True, type=int, help="Measured page count matching each --path")
    parser.add_argument("--title", action="append", default=[], help="Optional title matching each --path")
    parser.add_argument("--report", type=Path, help="Optional JSON report")
    args = parser.parse_args()
    if len(args.path) != len(args.pages):
        parser.error("--path and --pages must have the same count")
    if len(args.title) > len(args.path):
        parser.error("--title cannot be supplied more times than --path")
    titles = list(args.title) + [""] * (len(args.path) - len(args.title))
    results = [finalize(path, pages, titles[index]) for index, (path, pages) in enumerate(zip(args.path, args.pages))]
    payload = {"results": results}
    if args.report:
        args.report.parent.mkdir(parents=True, exist_ok=True)
        args.report.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(payload, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
