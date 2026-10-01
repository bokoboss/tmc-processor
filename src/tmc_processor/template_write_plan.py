"""Resolve native-template writes before choosing an Excel transport.

The plan starts at the post-Review export payload. It does not derive traffic
movements or change Mapping, Analyze, or Review decisions.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime, time
from numbers import Number
from pathlib import Path
import re
from typing import Any

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.cell import get_column_letter, range_boundaries

from .diagram import MOVEMENT_CODES
from .report_template import (
    _column_for_key,
    _mapped_label_value,
    _metadata_values,
    _normalise_time_label,
    movement_values_by_code,
)


@dataclass(frozen=True)
class CellWrite:
    sheet: str
    cell: str
    value: Any
    source: str
    order: int
    force: bool = False

    @property
    def value_type(self) -> str:
        value = self.value
        if value is None or value is pd.NA or value is pd.NaT:
            return "blank"
        try:
            if pd.isna(value):
                return "blank"
        except (TypeError, ValueError):
            pass
        if isinstance(value, bool):
            return "boolean"
        if isinstance(value, Number):
            return "number"
        if isinstance(value, (date, datetime, time, pd.Timestamp)):
            return "date_time"
        return "formula" if isinstance(value, str) and value.startswith("=") else "text"


@dataclass(frozen=True)
class DiagramRow:
    row: int
    code: str
    total_formula: str
    pm_formula: str
    am_formula: str
    total_cache: float
    pm_cache: float
    am_cache: float


@dataclass(frozen=True)
class PeakFormulaRow:
    period: str
    start: str
    end: str
    worksheet_row: int
    helper_cell: str
    helper_value: int


@dataclass(frozen=True)
class PeakBindingMatch:
    period: str
    interval: str
    matched: bool
    worksheet_row: int | None
    helper_cell: str | None
    helper_value: int | None


def _match_peak_rows(
    time_labels: dict[int, Any],
    hourly_map: dict[str, Any],
    helper_column: str,
    metadata: dict[str, Any],
) -> tuple[PeakBindingMatch, ...]:
    first, last = int(hourly_map["first_data_row"]), int(hourly_map["last_data_row"])
    matches: list[PeakBindingMatch] = []
    for period in ("am", "pm"):
        start = _normalise_time_label(metadata.get(f"{period}_peak_start"))
        end = _normalise_time_label(metadata.get(f"{period}_peak_end"))
        interval = f"{start}-{end}" if start and end else ""
        rows = [
            row for row in range(first, last + 1)
            if interval and _normalise_time_label(time_labels[row]) == interval
        ]
        row = rows[0] if len(rows) == 1 else None
        matches.append(PeakBindingMatch(
            period, interval, row is not None, row,
            f"{helper_column}{row}" if row is not None else None,
            row - int(hourly_map["header_row"]) + 1 if row is not None else None,
        ))
    return tuple(matches)


def preflight_template_peak_binding(
    template_path: str | Path,
    template_map: dict[str, Any],
    metadata: dict[str, Any],
) -> tuple[PeakBindingMatch, ...]:
    """Read the native Summary time rows without changing the workbook."""
    hourly_map = template_map["hourly_movement_table"]
    helper_column = str(template_map["movement_diagram_cells"]["movement_value_rows"]["helper_column"])
    workbook = load_workbook(template_path, read_only=True, data_only=False)
    try:
        summary = workbook[str(template_map.get("template_sheet") or "Summary")]
        if helper_column != "U" or [summary[f"{helper_column}{row}"].value for row in range(9, 23)] != list(range(1, 15)):
            raise ValueError("Authoritative Summary!U9:U22 helper sequence must be 1..14.")
        labels = {
            row: summary[f"{hourly_map['time_column']}{row}"].value
            for row in range(int(hourly_map["first_data_row"]), int(hourly_map["last_data_row"]) + 1)
        }
        return _match_peak_rows(labels, hourly_map, helper_column, metadata)
    finally:
        workbook.close()


@dataclass(frozen=True)
class FormulaBinding:
    period: str
    code: str
    cell: str
    old_formula: str
    new_formula: str


@dataclass(frozen=True)
class TemplateWritePlan:
    template_sheet: str
    summary_writes: tuple[CellWrite, ...]
    support_sheets: tuple[tuple[str, pd.DataFrame], ...]
    diagram_sheet_name: str
    diagram_rows: tuple[DiagramRow, ...]
    chart_caches: tuple[tuple[str, tuple[Any, ...], tuple[Any, ...]], ...]
    preserved_formulas: tuple[tuple[str, str], ...]
    skipped_formula_writes: tuple[tuple[str, str], ...]
    peak_formula_rows: tuple[PeakFormulaRow, ...]
    formula_bindings: tuple[FormulaBinding, ...]

    @property
    def ordered_writes(self) -> tuple[CellWrite, ...]:
        """Legacy COM order, including every support and Diagram_Data cell."""
        result: list[CellWrite] = []

        def append(sheet: str, cell: str, value: Any, source: str, force: bool = False) -> None:
            result.append(CellWrite(sheet, cell, value, source, len(result) + 1, force))

        for sheet, frame in self.support_sheets:
            if not len(frame.columns):
                append(sheet, "A1", "", f"support.{sheet}.empty")
                continue
            for col, name in enumerate(frame.columns, 1):
                append(sheet, f"{get_column_letter(col)}1", str(name), f"support.{sheet}.header.{name}")
            for row_number, values in enumerate(frame.itertuples(index=False, name=None), 2):
                for col, value in enumerate(values, 1):
                    append(sheet, f"{get_column_letter(col)}{row_number}", value, f"support.{sheet}.{row_number}.{col}")
        for col, value in enumerate(("movement_code", "total_pcu", "pm_peak_pcu", "am_peak_pcu"), 1):
            append(self.diagram_sheet_name, f"{get_column_letter(col)}1", value, "diagram.header")
        for item in self.diagram_rows:
            for col, value in enumerate((item.code, item.total_formula, item.pm_formula, item.am_formula), 1):
                append(self.diagram_sheet_name, f"{get_column_letter(col)}{item.row}", value, f"diagram.{item.code}.{col}")
        for write in self.summary_writes:
            append(write.sheet, write.cell, write.value, write.source, write.force)
        for binding in self.formula_bindings:
            append(self.template_sheet, binding.cell, binding.new_formula, f"formula_binding.{binding.period}.{binding.code}", True)
        return tuple(result)


def _in_ranges(cell: str, ranges: list[str]) -> bool:
    column, row, _, _ = range_boundaries(cell)
    for area in ranges:
        first_col, first_row, last_col, last_row = range_boundaries(str(area).split("!")[-1].replace("$", ""))
        if first_col <= column <= last_col and first_row <= row <= last_row:
            return True
    return False


def _range_cells(area: str) -> list[str]:
    first_col, first_row, last_col, last_row = range_boundaries(area.replace("$", ""))
    return [
        f"{get_column_letter(column)}{row}"
        for row in range(first_row, last_row + 1)
        for column in range(first_col, last_col + 1)
    ]


def diagram_data_formula(kind: str, code_ref: str, hourly_rows: int, hourly_columns: int) -> str:
    last = get_column_letter(max(hourly_columns, 2))
    data = f"'Hourly_Movement_PCU'!$B$2:${last}${hourly_rows}"
    times = f"'Hourly_Movement_PCU'!$A$2:$A${hourly_rows}"
    headers = f"'Hourly_Movement_PCU'!$B$1:${last}$1"
    match = f"MATCH({code_ref},{headers},0)"
    if kind == "total":
        return f"=IFERROR(INDEX({data},ROWS({times}),{match}),0)"
    peak = "'Peak_PHF'!$H$2" if kind == "pm" else "'Peak_PHF'!$D$2"
    key = f'IF(ISNUMBER({peak}),TEXT({peak},"hh:mm"),LEFT({peak},5))&"*"'
    return f"=IFERROR(INDEX({data},MATCH({key},{times},0),{match}),0)"


def resolve_template_write_plan(
    template_path: str | Path,
    template_map: dict[str, Any],
    report_data: dict[str, Any],
    metadata: dict[str, Any],
    chart_source_data: dict[str, Any],
) -> TemplateWritePlan:
    """Resolve the COM destination order into transport-independent cell writes."""

    template_sheet = str(template_map.get("template_sheet") or "Summary")
    workbook = load_workbook(template_path, read_only=True, data_only=False)
    try:
        summary = workbook[template_sheet]
        formulas = {
            cell.coordinate: str(cell.value)
            for row in summary.iter_rows()
            for cell in row
            if cell.data_type == "f"
        }
        original_times = {
            int(table["first_data_row"]): {
                row: summary[f"{table['time_column']}{row}"].value
                for row in range(int(table["first_data_row"]), int(table["total_row"]) + 1)
            }
            for table in (template_map["hourly_movement_table"], template_map["hourly_vehicle_class_table"])
        }
        helper_column = str(template_map["movement_diagram_cells"]["movement_value_rows"]["helper_column"])
        if helper_column != "U" or [summary[f"{helper_column}{row}"].value for row in range(9, 23)] != list(range(1, 15)):
            raise ValueError("Authoritative Summary!U9:U22 helper sequence must be 1..14.")
    finally:
        workbook.close()

    writable = list(template_map.get("writable_ranges") or [])
    overwritten = list(template_map.get("formula_overwrite_ranges") or [])
    protected = list(template_map.get("protected_formula_ranges") or [])
    writes: list[CellWrite] = []
    skipped: list[tuple[str, str]] = []

    def add(cell: str | None, value: Any, source: str, *, force: bool = False) -> None:
        if not cell:
            return
        if not force and not (_in_ranges(cell, writable) or _in_ranges(cell, overwritten)):
            if cell in formulas and _in_ranges(cell, protected):
                skipped.append((cell, source))
                return
            raise ValueError(f"Unmapped native-template write: {template_sheet}!{cell} ({source}).")
        if cell in formulas and not (force or _in_ranges(cell, overwritten)):
            skipped.append((cell, source))
            return
        writes.append(CellWrite(template_sheet, cell, value, source, len(writes) + 1, force))

    values = _metadata_values(metadata)
    for key, info in template_map.get("metadata_cells", {}).items():
        add(info.get("value_cell") or info.get("cell"), values.get(key, ""), f"metadata.{key}")
    movement = template_map.get("movement_diagram_cells", {})
    add(movement.get("diagram_title", {}).get("cell"), values.get("survey_point", ""), "movement.diagram_title")
    add(movement.get("diagram_date", {}).get("cell"), values.get("survey_date", ""), "movement.diagram_date")
    for key, info in movement.get("direction_labels", {}).items():
        add(info.get("cell"), _mapped_label_value(metadata, key), f"movement.direction_labels.{key}")
    for key, info in movement.get("road_labels", {}).items():
        add(info.get("cell"), _mapped_label_value(metadata, key), f"movement.road_labels.{key}")
    add(movement.get("caption", {}).get("cell"), metadata.get("caption_text") or "", "movement.caption")

    def peak_label(period: str) -> str:
        start = metadata.get(f"{period}_peak_start")
        end = metadata.get(f"{period}_peak_end")
        return f"{start}-{end}" if start and end else "-"

    provenance = str(metadata.get("peak_selection_source") or "").strip() or "-"
    add("A3", f"Peak used for export: AM {peak_label('am')}; PM {peak_label('pm')} ({provenance})", "effective_peak_label", force=True)

    def add_table(table: dict[str, Any], frame: pd.DataFrame, source_name: str) -> None:
        first = int(table["first_data_row"])
        total = int(table["total_row"])
        time_column = str(table["time_column"])
        labels = original_times[first]
        source_time = _column_for_key(frame, "time")
        by_label: dict[str, pd.Series] = {}
        source_total: pd.Series | None = None
        if source_time:
            for _, source_row in frame.iterrows():
                label = _normalise_time_label(source_row[source_time])
                if label in {"รวม", "total"}:
                    source_total = source_row
                elif label:
                    by_label[label] = source_row
        for key, column in table.get("columns", {}).items():
            source_column = _column_for_key(frame, key)
            add(f"{column}{table['header_row']}", source_column or key, f"{source_name}.header.{key}")
            for row in range(first, total + 1):
                add(f"{column}{row}", "", f"{source_name}.clear.{key}")
            for row in range(first, total + 1):
                source_row = source_total if row == total else by_label.get(_normalise_time_label(labels[row]))
                if key == "time":
                    value = labels[row]
                elif source_row is not None and source_column is not None:
                    value = source_row[source_column]
                else:
                    value = 0
                add(f"{column}{row}", value, f"{source_name}.{key}.{row}")

    hourly = report_data.get("hourly_movement_pcu", pd.DataFrame())
    add_table(template_map["hourly_movement_table"], hourly, "hourly_movement_pcu")
    hourly_map = template_map["hourly_movement_table"]
    first, last = int(hourly_map["first_data_row"]), int(hourly_map["last_data_row"])
    time_labels = original_times[first]
    peak_rows: list[PeakFormulaRow] = []
    for match in _match_peak_rows(time_labels, hourly_map, helper_column, metadata):
        if not match.interval:
            raise ValueError(f"Effective {match.period.upper()} Peak interval is required for native formula binding.")
        if not match.matched or match.worksheet_row is None or match.helper_cell is None or match.helper_value is None:
            raise ValueError(f"Effective {match.period.upper()} Peak {match.interval!r} must match exactly one Summary time row.")
        start, end = match.interval.split("-", 1)
        peak_rows.append(PeakFormulaRow(match.period, start, end, match.worksheet_row, match.helper_cell, match.helper_value))

    movement_map = template_map["movement_diagram_cells"]
    movement_columns = [column for key, column in hourly_map["columns"].items() if key not in {"time", "Total"}]
    lookup_range = f"${movement_columns[0]}${hourly_map['header_row']}:${movement_columns[-1]}${hourly_map['total_row']}"
    bindings: list[FormulaBinding] = []
    mapped_codes: set[str] = set()
    for approach in movement_map["approach_tables"].values():
        for code in approach["movement_codes"]:
            if code in mapped_codes:
                raise ValueError(f"Duplicate native movement formula mapping for {code}.")
            mapped_codes.add(code)
            header_cell = approach["header_cells"][code]
            for peak in peak_rows:
                cell = approach[f"{peak.period}_peak_hour_cells"][code]
                old = formulas.get(cell, "")
                pattern = rf"=HLOOKUP\({re.escape(header_cell)},{re.escape(lookup_range)},\${helper_column}\$(\d+),FALSE\)"
                match = re.fullmatch(pattern, old)
                if match is None or not first <= int(match.group(1)) <= last:
                    raise ValueError(f"Unknown native HLOOKUP formula at Summary!{cell}: {old!r}.")
                new = old[:match.start(1)] + str(peak.worksheet_row) + old[match.end(1):]
                bindings.append(FormulaBinding(peak.period, code, cell, old, new))
            total_cell = approach["total_12_hour_cells"][code]
            total = f"=HLOOKUP({header_cell},{lookup_range},${helper_column}${hourly_map['total_row']},FALSE)"
            if formulas.get(total_cell) != total:
                raise ValueError(f"Unknown native total HLOOKUP formula at Summary!{total_cell}.")
    if mapped_codes != set(MOVEMENT_CODES) or len(bindings) != 32 or len({item.cell for item in bindings}) != 32:
        raise ValueError("Native Peak formula map must contain 16 unique movements and 32 AM/PM cells.")
    support = report_data.get("sheets") or {}
    vehicle = report_data.get("hourly_vehicle_class")
    if vehicle is None:
        vehicle = support.get("Hourly_Vehicle_Class", pd.DataFrame())
    add_table(template_map["hourly_vehicle_class_table"], vehicle, "hourly_vehicle_class")

    chart_sources = template_map.get("chart_anchors", {})
    for source_key, data_key in (("native_hourly_chart_source", "hourly_pcu"), ("native_vehicle_composition_chart_source", "vehicle_composition")):
        source = chart_sources.get(source_key, {})
        data = chart_source_data.get(data_key, {})
        for kind in ("categories", "values"):
            area = source.get(kind)
            if not area:
                continue
            refs = _range_cells(area)
            series = list(data.get(kind, []))
            if len(series) > len(refs):
                raise ValueError(f"Chart data exceeds mapped range {area}.")
            for index, ref in enumerate(refs):
                add(ref, series[index] if index < len(series) else "", f"chart.{data_key}.{kind}.{index}")

    codes = tuple(report_data.get("diagram_movement_codes") or MOVEMENT_CODES)
    diagram_values = movement_values_by_code(hourly, list(codes), metadata)
    diagram_rows = tuple(
        DiagramRow(
            row,
            str(code),
            diagram_data_formula("total", f"$A{row}", max(len(hourly) + 1, 2), max(len(hourly.columns), 2)),
            diagram_data_formula("pm", f"$A{row}", max(len(hourly) + 1, 2), max(len(hourly.columns), 2)),
            diagram_data_formula("am", f"$A{row}", max(len(hourly) + 1, 2), max(len(hourly.columns), 2)),
            float(diagram_values[str(code)]["total_12_hour"]),
            float(diagram_values[str(code)]["pm_peak"]),
            float(diagram_values[str(code)]["am_peak"]),
        )
        for row, code in enumerate(codes, 2)
    )
    rebound_cells = {binding.cell for binding in bindings}
    preserved = tuple(sorted((cell, formula) for cell, formula in formulas.items() if not _in_ranges(cell, overwritten) and cell not in rebound_cells))
    caches = tuple(
        (key, tuple(chart_source_data.get(key, {}).get("categories", ())), tuple(chart_source_data.get(key, {}).get("values", ())))
        for key in ("hourly_pcu", "vehicle_composition")
    )
    return TemplateWritePlan(
        template_sheet,
        tuple(writes),
        tuple((name, frame.copy(deep=True)) for name, frame in support.items()),
        str(report_data.get("diagram_data_sheet_name") or "Diagram_Data"),
        diagram_rows,
        caches,
        preserved,
        tuple(skipped),
        tuple(peak_rows),
        tuple(bindings),
    )
