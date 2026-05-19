"""Read-only validation helpers for Excel report layout templates."""

from __future__ import annotations

import re
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from openpyxl.utils.cell import range_boundaries


EXTERNAL_LINK_RE = re.compile(r"\[[^\]]+\]")
QUOTED_SHEET_RE = re.compile(r"'((?:[^']|'')+)'!")
UNQUOTED_SHEET_RE = re.compile(r"(?<![\w\]])([A-Za-z_][A-Za-z0-9_ .]*)!")
CELL_REF_RE = re.compile(r"^\$?[A-Z]{1,3}\$?\d+$")
RANGE_REF_RE = re.compile(r"^\$?[A-Z]{1,3}\$?\d+:\$?[A-Z]{1,3}\$?\d+$")

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CHART_NS_URI = "http://schemas.openxmlformats.org/drawingml/2006/chart"
MAIN = {"m": MAIN_NS}
CHART = {"c": CHART_NS_URI}


@dataclass(frozen=True)
class FormulaIssue:
    sheet: str
    cell: str
    formula: str
    reason: str


@dataclass(frozen=True)
class ChartIssue:
    chart_part: str
    formula: str
    reason: str


@dataclass(frozen=True)
class MissingSheetIssue:
    sheet: str
    cell: str
    formula: str
    referenced_sheet: str


@dataclass(frozen=True)
class FormulaCell:
    sheet: str
    cell: str
    formula: str


@dataclass(frozen=True)
class MappedFormulaCell:
    sheet: str
    cell: str
    formula: str
    mapped_range: str


@dataclass(frozen=True)
class TemplateFormulaAudit:
    formula_cells: tuple[FormulaCell, ...]
    external_links: tuple[FormulaIssue, ...]
    ref_errors: tuple[FormulaIssue, ...]
    missing_sheet_issues: tuple[MissingSheetIssue, ...]
    mapped_formula_cells: tuple[MappedFormulaCell, ...]


@dataclass(frozen=True)
class TemplateAuditResult:
    formula_issues: tuple[FormulaIssue, ...]
    chart_issues: tuple[ChartIssue, ...]
    missing_sheet_issues: tuple[MissingSheetIssue, ...]

    @property
    def has_warnings(self) -> bool:
        return bool(self.formula_issues or self.chart_issues or self.missing_sheet_issues)

    @property
    def is_safe_for_template_export(self) -> bool:
        """True when no unsafe formula/chart references were detected."""

        return not self.has_warnings

    def warning_lines(self) -> list[str]:
        lines: list[str] = []
        for issue in self.formula_issues:
            lines.append(f"{issue.sheet}!{issue.cell}: {issue.reason}: {issue.formula}")
        for issue in self.chart_issues:
            lines.append(f"{issue.chart_part}: {issue.reason}: {issue.formula}")
        for issue in self.missing_sheet_issues:
            lines.append(
                f"{issue.sheet}!{issue.cell}: references missing sheet "
                f"{issue.referenced_sheet!r}: {issue.formula}"
            )
        return lines


@dataclass(frozen=True)
class TemplateUseValidation:
    """Decision helper for template-driven export.

    Broken native charts are warnings because export code should ignore existing
    chart objects and insert generated PNGs at mapped anchors instead.
    Unsafe worksheet formulas and missing sheet references are blocking because
    they can trigger Excel repair prompts or stale external-link prompts.
    """

    can_use_template_cells: bool
    blocking_warnings: tuple[str, ...]
    non_blocking_warnings: tuple[str, ...]

    @property
    def should_fallback_to_generated_report(self) -> bool:
        return not self.can_use_template_cells


def _normalise_sheet_name(name: str) -> str:
    return name.replace("''", "'").strip()


def _formula_sheet_references(formula: str) -> set[str]:
    references: set[str] = set()
    for match in QUOTED_SHEET_RE.finditer(formula):
        raw = match.group(1)
        if EXTERNAL_LINK_RE.search(raw):
            continue
        references.add(_normalise_sheet_name(raw))
    for match in UNQUOTED_SHEET_RE.finditer(formula):
        raw = match.group(1).strip()
        if EXTERNAL_LINK_RE.search(raw):
            continue
        if CELL_REF_RE.match(raw) or RANGE_REF_RE.match(raw):
            continue
        references.add(raw)
    return references


def sheet_parts(path: Path) -> dict[str, str]:
    with zipfile.ZipFile(path) as archive:
        workbook_root = ET.fromstring(archive.read("xl/workbook.xml"))
        rels_root = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        rel_targets = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels_root}
        sheets: dict[str, str] = {}
        for sheet in workbook_root.findall(".//m:sheet", MAIN):
            name = sheet.attrib["name"]
            rel_id = sheet.attrib[f"{{{REL_NS}}}id"]
            target = rel_targets[rel_id]
            sheets[name] = "xl/" + target.lstrip("/")
        return sheets


def iter_formula_cells(path: Path) -> tuple[set[str], list[tuple[str, str, str]]]:
    formulas: list[tuple[str, str, str]] = []
    parts = sheet_parts(path)
    with zipfile.ZipFile(path) as archive:
        for sheet_name, sheet_part in parts.items():
            root = ET.fromstring(archive.read(sheet_part))
            for cell in root.findall(".//m:c", MAIN):
                formula_node = cell.find("m:f", MAIN)
                if formula_node is not None and formula_node.text:
                    formulas.append((sheet_name, cell.attrib.get("r", ""), formula_node.text))
    return set(parts), formulas


def chart_formulas(path: Path) -> list[tuple[str, str]]:
    formulas: list[tuple[str, str]] = []
    with zipfile.ZipFile(path) as archive:
        chart_parts = sorted(
            name for name in archive.namelist() if name.startswith("xl/charts/") and name.endswith(".xml")
        )
        for chart_part in chart_parts:
            root = ET.fromstring(archive.read(chart_part))
            for formula_node in root.findall(".//c:f", CHART):
                if formula_node.text:
                    formulas.append((chart_part, formula_node.text))
    return formulas


def _strip_sheet_reference(reference: str) -> tuple[str | None, str]:
    text = str(reference).strip()
    if "!" not in text:
        return None, text.replace("$", "")
    sheet, cell_range = text.rsplit("!", 1)
    return sheet.strip("'"), cell_range.replace("$", "")


def _cell_in_range(cell: str, cell_range: str) -> bool:
    _, cleaned_range = _strip_sheet_reference(cell_range)
    min_col, min_row, max_col, max_row = range_boundaries(cleaned_range)
    cell_col, cell_row, _, _ = range_boundaries(cell)
    return min_col <= cell_col <= max_col and min_row <= cell_row <= max_row


def _iter_map_ranges(value: Any) -> list[str]:
    ranges: list[str] = []
    if isinstance(value, dict):
        for key, child in value.items():
            if key == "protected_formula_ranges":
                continue
            if isinstance(child, str) and (
                key == "cell" or key.endswith("_cell") or key == "range" or key.endswith("_range")
            ):
                ranges.append(child)
            else:
                ranges.extend(_iter_map_ranges(child))
    elif isinstance(value, list):
        for child in value:
            if isinstance(child, str):
                ranges.append(child)
            else:
                ranges.extend(_iter_map_ranges(child))
    return ranges


def _mapped_write_ranges(mapping: dict[str, Any]) -> list[str]:
    ranges: list[str] = []
    for key in ("writable_ranges", "formula_overwrite_ranges"):
        ranges.extend(_iter_map_ranges(mapping.get(key, [])))
    for key in (
        "metadata_cells",
        "movement_diagram_cells",
        "summary_box_cells",
        "hourly_movement_table",
        "hourly_vehicle_class_table",
    ):
        ranges.extend(_iter_map_ranges(mapping.get(key, {})))
    anchors = mapping.get("chart_anchors", {})
    for key in ("native_hourly_chart_source", "native_vehicle_composition_chart_source"):
        native_source = anchors.get(key, {})
        ranges.extend(str(value) for value in native_source.values() if isinstance(value, str))
    return ranges


def mapped_formula_cells(
    path: str | Path,
    mapping: dict[str, Any],
    sheet_name: str = "Summary",
) -> tuple[MappedFormulaCell, ...]:
    _, formulas = iter_formula_cells(Path(path))
    mapped_ranges = _mapped_write_ranges(mapping)
    issues: list[MappedFormulaCell] = []
    for sheet, cell, formula in formulas:
        if sheet != sheet_name:
            continue
        for mapped_range in mapped_ranges:
            mapped_sheet, _ = _strip_sheet_reference(mapped_range)
            if mapped_sheet and mapped_sheet != sheet:
                continue
            try:
                if _cell_in_range(cell, mapped_range):
                    issues.append(MappedFormulaCell(sheet, cell, formula, mapped_range))
                    break
            except ValueError:
                continue
    return tuple(issues)


def audit_template_formulas(
    path: str | Path,
    mapping: dict[str, Any] | None = None,
    sheet_name: str = "Summary",
) -> TemplateFormulaAudit:
    template_path = Path(path)
    sheet_names, formulas = iter_formula_cells(template_path)
    formula_cells = tuple(
        FormulaCell(sheet, cell, formula)
        for sheet, cell, formula in formulas
        if sheet == sheet_name
    )
    external_links: list[FormulaIssue] = []
    ref_errors: list[FormulaIssue] = []
    missing_sheet_issues: list[MissingSheetIssue] = []

    for sheet, cell, formula in formulas:
        if sheet != sheet_name:
            continue
        if "#REF!" in formula:
            ref_errors.append(FormulaIssue(sheet, cell, formula, "contains #REF!"))
        if EXTERNAL_LINK_RE.search(formula):
            external_links.append(FormulaIssue(sheet, cell, formula, "contains external workbook link"))
        for referenced_sheet in sorted(_formula_sheet_references(formula)):
            if referenced_sheet not in sheet_names:
                missing_sheet_issues.append(MissingSheetIssue(sheet, cell, formula, referenced_sheet))

    mapped = mapped_formula_cells(template_path, mapping, sheet_name) if mapping else ()
    return TemplateFormulaAudit(
        formula_cells=formula_cells,
        external_links=tuple(external_links),
        ref_errors=tuple(ref_errors),
        missing_sheet_issues=tuple(missing_sheet_issues),
        mapped_formula_cells=mapped,
    )


def audit_template(path: str | Path) -> TemplateAuditResult:
    template_path = Path(path)
    sheet_names, formulas = iter_formula_cells(template_path)
    formula_issues: list[FormulaIssue] = []
    missing_sheet_issues: list[MissingSheetIssue] = []

    for sheet, cell, formula in formulas:
        if "#REF!" in formula:
            formula_issues.append(FormulaIssue(sheet, cell, formula, "contains #REF!"))
        if EXTERNAL_LINK_RE.search(formula):
            formula_issues.append(FormulaIssue(sheet, cell, formula, "contains external workbook link"))
        for referenced_sheet in sorted(_formula_sheet_references(formula)):
            if referenced_sheet not in sheet_names:
                missing_sheet_issues.append(MissingSheetIssue(sheet, cell, formula, referenced_sheet))

    chart_issues = []
    for chart_part, formula in chart_formulas(template_path):
        if "#REF!" in formula:
            chart_issues.append(ChartIssue(chart_part, formula, "contains #REF!"))
        if EXTERNAL_LINK_RE.search(formula):
            chart_issues.append(ChartIssue(chart_part, formula, "contains external workbook link"))
    return TemplateAuditResult(tuple(formula_issues), chart_issues, tuple(missing_sheet_issues))


def validate_template_before_export(path: str | Path) -> TemplateUseValidation:
    """Validate a template before writing mapped cells into it.

    Export callers should only write cells listed in the JSON map, ignore
    existing native chart objects, and place generated PNG charts at mapped PNG
    anchors. If this validation returns ``should_fallback_to_generated_report``,
    callers should use the generated report workbook path instead of the layout
    template.
    """

    result = audit_template(path)
    blocking = [
        *(
            f"{issue.sheet}!{issue.cell}: {issue.reason}: {issue.formula}"
            for issue in result.formula_issues
        ),
        *(
            f"{issue.sheet}!{issue.cell}: references missing sheet "
            f"{issue.referenced_sheet!r}: {issue.formula}"
            for issue in result.missing_sheet_issues
        ),
    ]
    chart_warnings = tuple(
        f"{issue.chart_part}: {issue.reason}: {issue.formula}" for issue in result.chart_issues
    )
    return TemplateUseValidation(
        can_use_template_cells=not blocking,
        blocking_warnings=tuple(blocking),
        non_blocking_warnings=chart_warnings,
    )
