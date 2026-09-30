"""Inspect XLSX package parts and run a one-cell OOXML preservation proof.

This is an isolated qualification utility. It is not wired into application exports.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import posixpath
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Iterable
from xml.dom import Node, minidom
from zipfile import ZipFile

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
DOC_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
XML_NS = "http://www.w3.org/XML/1998/namespace"
PRESERVE_PREFIXES = (
    "xl/charts/",
    "xl/drawings/",
    "xl/media/",
    "xl/embeddings/",
    "xl/theme/",
)


@dataclass(frozen=True)
class PackagePart:
    path: str
    size: int
    sha256: str
    category: str


def worksheet_part_for_name(archive: ZipFile, sheet_name: str) -> str:
    """Resolve a worksheet name through workbook relationships."""
    workbook = minidom.parseString(archive.read("xl/workbook.xml"))
    rels = minidom.parseString(archive.read("xl/_rels/workbook.xml.rels"))
    targets: dict[str, str] = {}
    for rel in rels.getElementsByTagNameNS(PKG_REL_NS, "Relationship"):
        if rel.getAttribute("TargetMode").casefold() == "external":
            continue
        targets[rel.getAttribute("Id")] = rel.getAttribute("Target")

    for sheet in workbook.getElementsByTagNameNS(MAIN_NS, "sheet"):
        if sheet.getAttribute("name") != sheet_name:
            continue
        relationship_id = sheet.getAttributeNS(DOC_REL_NS, "id")
        target = targets.get(relationship_id)
        if not target:
            raise ValueError(f"Worksheet relationship is missing for {sheet_name!r}.")
        part = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join("xl", target))
        if part not in archive.namelist():
            raise ValueError(f"Worksheet part is absent from the package: {part}")
        return part
    raise ValueError(f"Worksheet not found: {sheet_name!r}")


def inventory_package(
    package_path: str | Path,
    *,
    writable_parts: Iterable[str] = (),
    structural_parts: Iterable[str] = (),
) -> tuple[PackagePart, ...]:
    """Return a deterministic, sorted inventory of every file part in an XLSX."""
    writable = set(writable_parts)
    structural = set(structural_parts)
    if writable & structural:
        raise ValueError("A part cannot be both writable and structural.")
    records: list[PackagePart] = []
    with ZipFile(package_path) as archive:
        for name in sorted(archive.namelist()):
            content = archive.read(name)
            if name in writable:
                category = "expected writable"
            elif name in structural:
                category = "expected structural change"
            else:
                category = "must remain byte-identical"
            records.append(PackagePart(name, len(content), hashlib.sha256(content).hexdigest(), category))
    return tuple(records)


def _inline_string_cell_xml(sheet_xml: bytes, cell_ref: str, value: str) -> bytes:
    """Namespace-aware update of one existing worksheet cell to an inline string."""
    document = minidom.parseString(sheet_xml)
    cells = [
        cell
        for cell in document.getElementsByTagNameNS(MAIN_NS, "c")
        if cell.getAttribute("r") == cell_ref
    ]
    if len(cells) != 1:
        raise ValueError(f"Expected exactly one existing {cell_ref} cell; found {len(cells)}.")
    cell = cells[0]
    if any(
        child.nodeType == Node.ELEMENT_NODE and child.namespaceURI == MAIN_NS and child.localName == "f"
        for child in cell.childNodes
    ):
        raise ValueError(f"Refusing to replace a formula in {cell_ref}.")

    style = cell.getAttribute("s")
    for child in list(cell.childNodes):
        cell.removeChild(child)
        child.unlink()
    if cell.hasAttribute("t"):
        cell.removeAttribute("t")
    cell.setAttribute("t", "inlineStr")
    inline = document.createElementNS(MAIN_NS, "is")
    text_node = document.createElementNS(MAIN_NS, "t")
    if value[:1].isspace() or value[-1:].isspace():
        text_node.setAttributeNS(XML_NS, "xml:space", "preserve")
    text_node.appendChild(document.createTextNode(value))
    inline.appendChild(text_node)
    cell.appendChild(inline)
    if cell.getAttribute("s") != style:
        raise AssertionError("Cell style changed during the value update.")
    return document.toxml(encoding="utf-8")


def patch_one_cell(
    source_path: str | Path,
    output_path: str | Path,
    *,
    sheet_name: str,
    cell_ref: str,
    value: str,
) -> str:
    """Copy an XLSX and change one existing string cell without using Excel/openpyxl."""
    source = Path(source_path).resolve()
    output = Path(output_path).resolve()
    if source == output:
        raise ValueError("Output must not overwrite the source template.")
    output.parent.mkdir(parents=True, exist_ok=True)
    with ZipFile(source, "r") as original:
        sheet_part = worksheet_part_for_name(original, sheet_name)
        patched_sheet = _inline_string_cell_xml(original.read(sheet_part), cell_ref, value)
        with ZipFile(output, "w") as patched:
            patched.comment = original.comment
            for info in original.infolist():
                content = patched_sheet if info.filename == sheet_part else original.read(info.filename)
                patched.writestr(info, content)
    return sheet_part


def _json_records(records: Iterable[PackagePart]) -> list[dict[str, object]]:
    return [asdict(record) for record in records]


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("template", type=Path)
    parser.add_argument("--sheet", default="Summary")
    parser.add_argument("--cell", default="E5")
    parser.add_argument("--value", default="OOXML package patch proof")
    parser.add_argument("--output", type=Path, help="Write a one-cell proof workbook and report both inventories.")
    parser.add_argument("--compare-output", type=Path, help="Classify every template/output package part against mutation allowlists.")
    parser.add_argument("--expected-changed", nargs="*", default=[])
    parser.add_argument("--expected-added", nargs="*", default=[])
    parser.add_argument("--expected-removed", nargs="*", default=[])
    args = parser.parse_args()

    if args.compare_output:
        entries = compare_package_parts(
            args.template,
            args.compare_output,
            expected_changed=args.expected_changed,
            expected_added=args.expected_added,
            expected_removed=args.expected_removed,
        )
        report = {
            "template": str(args.template.resolve()),
            "output": str(args.compare_output.resolve()),
            "parts": _json_records(entries),
            "summary": summarize_package_diff(entries),
            "unexpected_parts": [entry.template_part_path or entry.output_part_path for entry in entries if entry.classification.startswith("UNEXPECTED_")],
        }
        print(json.dumps(report, ensure_ascii=True, indent=2))
        return

    with ZipFile(args.template) as archive:
        sheet_part = worksheet_part_for_name(archive, args.sheet)
    if args.output:
        changed_part = patch_one_cell(
            args.template, args.output, sheet_name=args.sheet, cell_ref=args.cell, value=args.value
        )
        report = {
            "template": str(args.template.resolve()),
            "output": str(args.output.resolve()),
            "changed_part": changed_part,
            "before": _json_records(inventory_package(args.template, writable_parts=(sheet_part,))),
            "after": _json_records(inventory_package(args.output, writable_parts=(sheet_part,))),
        }
    else:
        report = {
            "template": str(args.template.resolve()),
            "inventory": _json_records(inventory_package(args.template, writable_parts=(sheet_part,))),
        }
    print(json.dumps(report, ensure_ascii=True, indent=2))


@dataclass(frozen=True)
class PackageDiffEntry:
    template_part_path: str | None
    output_part_path: str | None
    template_sha256: str | None
    output_sha256: str | None
    classification: str


def compare_package_parts(
    template_path: str | Path,
    output_path: str | Path,
    *,
    expected_changed: Iterable[str] = (),
    expected_added: Iterable[str] = (),
    expected_removed: Iterable[str] = (),
) -> tuple[PackageDiffEntry, ...]:
    """Classify every package part against an explicit mutation allowlist."""
    changed = set(expected_changed)
    added = set(expected_added)
    removed = set(expected_removed)
    if changed & (added | removed) or added & removed:
        raise ValueError("Expected changed/added/removed part sets must be disjoint.")
    with ZipFile(template_path) as source, ZipFile(output_path) as output:
        before = {name: source.read(name) for name in source.namelist()}
        after = {name: output.read(name) for name in output.namelist()}
    entries = []
    for name in sorted(before.keys() | after.keys()):
        if name in before and name in after:
            same = before[name] == after[name]
            classification = "UNCHANGED_EXPECTED" if same else ("CHANGED_EXPECTED" if name in changed else "UNEXPECTED_CHANGE")
            entries.append(PackageDiffEntry(name, name, hashlib.sha256(before[name]).hexdigest(), hashlib.sha256(after[name]).hexdigest(), classification))
        elif name in before:
            classification = "REMOVED_EXPECTED" if name in removed else "UNEXPECTED_MISSING"
            entries.append(PackageDiffEntry(name, None, hashlib.sha256(before[name]).hexdigest(), None, classification))
        else:
            classification = "ADDED_EXPECTED" if name in added else "UNEXPECTED_ADDITION"
            entries.append(PackageDiffEntry(None, name, None, hashlib.sha256(after[name]).hexdigest(), classification))
    return tuple(entries)


def package_part_category(path: str) -> str:
    if path in {"xl/workbook.xml", "xl/calcChain.xml", "[Content_Types].xml", "_rels/.rels"}: return "workbook core"
    if path == "xl/_rels/workbook.xml.rels": return "workbook relationships"
    if path.startswith("xl/worksheets/_rels/"): return "worksheet relationships"
    if path.startswith("xl/worksheets/"): return "support worksheets" if path != "xl/worksheets/sheet1.xml" else "Summary worksheet"
    if path.startswith("xl/charts/"): return "charts"
    if path.startswith("xl/drawings/"): return "drawings"
    if path.startswith("xl/media/"): return "media"
    if path == "xl/styles.xml": return "styles"
    if path.startswith("xl/theme/"): return "theme"
    if path.startswith("docProps/"): return "properties"
    return "other"


def summarize_package_diff(entries: Iterable[PackageDiffEntry]) -> dict[str, dict[str, int]]:
    summary: dict[str, dict[str, int]] = {}
    for entry in entries:
        path = entry.output_part_path or entry.template_part_path or ""
        category = package_part_category(path)
        counts = summary.setdefault(category, {"unchanged": 0, "changed expected": 0, "added expected": 0, "removed expected": 0, "unexpected": 0})
        key = {"UNCHANGED_EXPECTED": "unchanged", "CHANGED_EXPECTED": "changed expected", "ADDED_EXPECTED": "added expected", "REMOVED_EXPECTED": "removed expected"}.get(entry.classification, "unexpected")
        counts[key] += 1
    return summary


if __name__ == "__main__":
    main()
