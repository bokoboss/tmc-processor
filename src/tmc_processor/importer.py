"""Excel importing and raw Thai TMC count-table parsing."""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import datetime, time
from pathlib import Path
from typing import BinaryIO, NamedTuple

import pandas as pd

from .time_utils import parse_time


THAI_DIRECTION_PREFIX = "\u0e17\u0e34\u0e28"
RAW_DIRECTION_PATTERN = re.compile(rf"^\s*(?:{THAI_DIRECTION_PREFIX})\s*([0-9]+(?:\+[0-9]+)*)\s*$", re.IGNORECASE)
SCAN_ROWS = 30
HEADER_LOOKBACK_ROWS = 6


VEHICLE_ALIASES = {
    "Bicy": [
        "\u0e23\u0e16\u0e08\u0e31\u0e01\u0e23\u0e22\u0e32\u0e19",
        "\u0e08\u0e31\u0e01\u0e23\u0e22\u0e32\u0e19",
        "\u0e23\u0e16\u0e08\u0e31\u0e01\u0e23\u0e22\u0e32\u0e19 2 \u0e25\u0e49\u0e2d",
        "\u0e23\u0e16\u0e08\u0e31\u0e01\u0e23\u0e22\u0e32\u0e19 3 \u0e25\u0e49\u0e2d",
    ],
    "MC": [
        "\u0e23\u0e16\u0e08\u0e31\u0e01\u0e23\u0e22\u0e32\u0e19\u0e22\u0e19\u0e15\u0e4c",
        "\u0e23\u0e16\u0e2a\u0e32\u0e21\u0e25\u0e49\u0e2d\u0e40\u0e04\u0e23\u0e37\u0e48\u0e2d\u0e07",
        "\u0e08\u0e22\u0e22.",
        "\u0e21\u0e2d\u0e40\u0e15\u0e2d\u0e23\u0e4c\u0e44\u0e0b\u0e04\u0e4c",
    ],
    "PC<7": [
        "\u0e23\u0e16\u0e22\u0e19\u0e15\u0e4c\u0e19\u0e31\u0e48\u0e07\u0e2a\u0e48\u0e27\u0e19\u0e1a\u0e38\u0e04\u0e04\u0e25\u0e44\u0e21\u0e48\u0e40\u0e01\u0e34\u0e19 7 \u0e04\u0e19",
        "\u0e44\u0e21\u0e48\u0e40\u0e01\u0e34\u0e19 7 \u0e04\u0e19",
    ],
    "PC>7": [
        "\u0e23\u0e16\u0e15\u0e39\u0e49",
        "\u0e23\u0e16\u0e22\u0e19\u0e15\u0e4c\u0e19\u0e31\u0e48\u0e07\u0e2a\u0e48\u0e27\u0e19\u0e1a\u0e38\u0e04\u0e04\u0e25\u0e40\u0e01\u0e34\u0e19 7 \u0e04\u0e19",
        "\u0e40\u0e01\u0e34\u0e19 7 \u0e04\u0e19",
    ],
    "LB": [
        "\u0e23\u0e16\u0e42\u0e14\u0e22\u0e2a\u0e32\u0e23\u0e02\u0e19\u0e32\u0e14\u0e40\u0e25\u0e47\u0e01",
        "2 \u0e41\u0e16\u0e27\u0e40\u0e25\u0e47\u0e01",
    ],
    "MB": ["\u0e23\u0e16\u0e42\u0e14\u0e22\u0e2a\u0e32\u0e23\u0e02\u0e19\u0e32\u0e14\u0e01\u0e25\u0e32\u0e07"],
    "HB": ["\u0e23\u0e16\u0e42\u0e14\u0e22\u0e2a\u0e32\u0e23\u0e02\u0e19\u0e32\u0e14\u0e43\u0e2b\u0e0d\u0e48"],
    "LT": [
        "\u0e23\u0e16\u0e1a\u0e23\u0e23\u0e17\u0e38\u0e01\u0e02\u0e19\u0e32\u0e14\u0e40\u0e25\u0e47\u0e01",
        "4 \u0e25\u0e49\u0e2d",
    ],
    "MT": [
        "\u0e23\u0e16\u0e1a\u0e23\u0e23\u0e17\u0e38\u0e01\u0e02\u0e19\u0e32\u0e14\u0e01\u0e25\u0e32\u0e07",
        "6 \u0e25\u0e49\u0e2d",
    ],
    "HT": [
        "\u0e23\u0e16\u0e1a\u0e23\u0e23\u0e17\u0e38\u0e01\u0e02\u0e19\u0e32\u0e14\u0e43\u0e2b\u0e0d\u0e48",
        "10 \u0e25\u0e49\u0e2d",
    ],
    "TR": ["\u0e23\u0e16\u0e1a\u0e23\u0e23\u0e17\u0e38\u0e01\u0e1e\u0e48\u0e27\u0e07"],
    "STR": ["\u0e23\u0e16\u0e1a\u0e23\u0e23\u0e17\u0e38\u0e01\u0e01\u0e36\u0e48\u0e07\u0e1e\u0e48\u0e27\u0e07"],
}

EXCLUDED_HEADER_TERMS = ["pcu", "pce", "total", "grand total", "\u0e23\u0e27\u0e21"]


@dataclass(frozen=True)
class DetectedSheet:
    raw_sheet: str
    raw_direction: str
    rows: int
    columns: int


@dataclass(frozen=True)
class ParsedSheetDebug:
    raw_sheet: str
    raw_direction: str
    first_data_row: int | None
    time_start_col: int | None
    time_end_col: int | None
    vehicle_class_columns: list[dict[str, object]]


class ParsedSheet(NamedTuple):
    data: pd.DataFrame
    debug: ParsedSheetDebug
    raw: pd.DataFrame


def extract_raw_direction(sheet_name: str) -> str | None:
    match = RAW_DIRECTION_PATTERN.match(str(sheet_name))
    if not match:
        return None
    return match.group(1)


def detect_raw_direction_sheet_names(excel_file: str | Path | BinaryIO) -> list[str]:
    workbook = pd.ExcelFile(excel_file)
    return [name for name in workbook.sheet_names if extract_raw_direction(name) is not None]


def _text(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def _compact(value: str) -> str:
    return re.sub(r"[\s_]+", "", value).lower()


def _is_time_like(value: object) -> bool:
    if isinstance(value, (int, float)):
        return False
    if isinstance(value, (datetime, time, pd.Timestamp)):
        return parse_time(value) is not None
    if not re.search(r"\d{1,2}[:.]\d{2}", _text(value)):
        return False
    return parse_time(value) is not None


def _find_interval_in_row(row: pd.Series) -> tuple[int, int] | None:
    time_cols = [idx for idx, value in row.items() if _is_time_like(value)]
    for start_col in time_cols:
        for end_col in time_cols:
            if start_col < end_col <= start_col + 4:
                return int(start_col), int(end_col)
    return None


def _locate_count_table(raw: pd.DataFrame) -> tuple[int | None, int | None, int | None]:
    max_rows = min(SCAN_ROWS, len(raw))
    for row_idx in range(max_rows):
        interval = _find_interval_in_row(raw.iloc[row_idx])
        if interval:
            return row_idx, interval[0], interval[1]
    return None, None, None


def _header_text(raw: pd.DataFrame, first_data_row: int, column: int) -> str:
    start = max(0, first_data_row - HEADER_LOOKBACK_ROWS)
    header = raw.iloc[start:first_data_row, :].copy()
    header = header.ffill(axis=1)
    parts = [_text(value) for value in header.iloc[:, column].tolist()]
    unique_parts = []
    for part in parts:
        if part and part not in unique_parts:
            unique_parts.append(part)
    return " ".join(unique_parts)


def _raw_header_has_text(raw: pd.DataFrame, first_data_row: int, column: int) -> bool:
    start = max(0, first_data_row - HEADER_LOOKBACK_ROWS)
    return any(_text(value) for value in raw.iloc[start:first_data_row, column].tolist())


def _match_vehicle_class(header: str) -> str | None:
    compact_header = _compact(header)
    if not compact_header:
        return None
    if any(_compact(term) in compact_header for term in EXCLUDED_HEADER_TERMS):
        return None

    if _compact("\u0e01\u0e36\u0e48\u0e07\u0e1e\u0e48\u0e27\u0e07") in compact_header:
        return "STR"
    if _compact("\u0e23\u0e16\u0e1a\u0e23\u0e23\u0e17\u0e38\u0e01") in compact_header and _compact("\u0e1e\u0e48\u0e27\u0e07") in compact_header:
        return "TR"
    if _compact("\u0e23\u0e16\u0e15\u0e39\u0e49") in compact_header:
        return "PC>7"
    if _compact("\u0e44\u0e21\u0e48\u0e40\u0e01\u0e34\u0e197\u0e04\u0e19") in compact_header:
        return "PC<7"
    if _compact("\u0e40\u0e01\u0e34\u0e197\u0e04\u0e19") in compact_header:
        return "PC>7"
    if _compact("\u0e23\u0e16\u0e08\u0e31\u0e01\u0e23\u0e22\u0e32\u0e19\u0e22\u0e19\u0e15\u0e4c") in compact_header:
        return "MC"
    if _compact("\u0e23\u0e16\u0e2a\u0e32\u0e21\u0e25\u0e49\u0e2d\u0e40\u0e04\u0e23\u0e37\u0e48\u0e2d\u0e07") in compact_header:
        return "MC"
    if _compact("\u0e23\u0e16\u0e42\u0e14\u0e22\u0e2a\u0e32\u0e23") in compact_header and _compact("\u0e02\u0e19\u0e32\u0e14\u0e43\u0e2b\u0e0d\u0e48") in compact_header:
        return "HB"
    if _compact("\u0e23\u0e16\u0e42\u0e14\u0e22\u0e2a\u0e32\u0e23") in compact_header and _compact("\u0e02\u0e19\u0e32\u0e14\u0e01\u0e25\u0e32\u0e07") in compact_header:
        return "MB"
    if _compact("\u0e23\u0e16\u0e42\u0e14\u0e22\u0e2a\u0e32\u0e23") in compact_header and _compact("\u0e02\u0e19\u0e32\u0e14\u0e40\u0e25\u0e47\u0e01") in compact_header:
        return "LB"
    if _compact("\u0e23\u0e16\u0e1a\u0e23\u0e23\u0e17\u0e38\u0e01") in compact_header and _compact("10\u0e25\u0e49\u0e2d") in compact_header:
        return "HT"
    if _compact("\u0e23\u0e16\u0e1a\u0e23\u0e23\u0e17\u0e38\u0e01") in compact_header and _compact("6\u0e25\u0e49\u0e2d") in compact_header:
        return "MT"
    if _compact("\u0e23\u0e16\u0e1a\u0e23\u0e23\u0e17\u0e38\u0e01") in compact_header and _compact("4\u0e25\u0e49\u0e2d") in compact_header:
        return "LT"

    for standard in ["STR", "TR", "PC>7", "PC<7", "MC", "Bicy", "HB", "MB", "LB", "HT", "MT", "LT"]:
        aliases = VEHICLE_ALIASES.get(standard, [])
        for alias in aliases:
            compact_alias = _compact(alias)
            if standard == "PC>7" and _compact("\u0e44\u0e21\u0e48\u0e40\u0e01\u0e34\u0e19 7 \u0e04\u0e19") in compact_header:
                continue
            if compact_alias and compact_alias in compact_header:
                return standard
    return None


def _detect_vehicle_columns(raw: pd.DataFrame, first_data_row: int, time_end_col: int) -> list[dict[str, object]]:
    detected: list[dict[str, object]] = []
    found_block = False
    for column in range(time_end_col + 1, raw.shape[1]):
        if found_block and not _raw_header_has_text(raw, first_data_row, column):
            break
        header = _header_text(raw, first_data_row, column)
        if not header and found_block:
            break
        vehicle_class = _match_vehicle_class(header)
        if vehicle_class:
            found_block = True
            detected.append({"column": column, "vehicle_class": vehicle_class, "header": header})
    return detected


def _numeric_count(value: object) -> float:
    if pd.isna(value) or _text(value) == "":
        return 0.0
    numeric = pd.to_numeric(value, errors="coerce")
    return 0.0 if pd.isna(numeric) else float(numeric)


def parse_raw_count_sheet(raw: pd.DataFrame, raw_sheet: str) -> ParsedSheet:
    raw_direction = extract_raw_direction(raw_sheet) or ""
    first_data_row, time_start_col, time_end_col = _locate_count_table(raw)
    if first_data_row is None or time_start_col is None or time_end_col is None:
        debug = ParsedSheetDebug(raw_sheet, raw_direction, None, None, None, [])
        return ParsedSheet(pd.DataFrame(columns=["raw_sheet", "raw_direction", "time_start", "time_end", "vehicle_class", "count"]), debug, raw)

    vehicle_columns = _detect_vehicle_columns(raw, first_data_row, time_end_col)
    rows = []
    blank_interval_rows = 0
    for row_idx in range(first_data_row, len(raw)):
        row = raw.iloc[row_idx]
        time_start = parse_time(row.iloc[time_start_col])
        time_end = parse_time(row.iloc[time_end_col])
        if time_start is None or time_end is None:
            blank_interval_rows += 1
            if blank_interval_rows >= 3:
                break
            continue
        blank_interval_rows = 0
        for vehicle in vehicle_columns:
            rows.append(
                {
                    "raw_sheet": raw_sheet,
                    "raw_direction": raw_direction,
                    "time_start": time_start,
                    "time_end": time_end,
                    "vehicle_class": vehicle["vehicle_class"],
                    "count": _numeric_count(row.iloc[int(vehicle["column"])]),
                }
            )

    data = pd.DataFrame(rows, columns=["raw_sheet", "raw_direction", "time_start", "time_end", "vehicle_class", "count"])
    debug_vehicle_columns = [
        {**vehicle, "column": int(vehicle["column"]) + 1}
        for vehicle in vehicle_columns
    ]
    debug = ParsedSheetDebug(raw_sheet, raw_direction, first_data_row + 1, time_start_col + 1, time_end_col + 1, debug_vehicle_columns)
    return ParsedSheet(data, debug, raw)


def load_detected_sheet_details(excel_file: str | Path | BinaryIO) -> dict[str, ParsedSheet]:
    sheet_names = detect_raw_direction_sheet_names(excel_file)
    if hasattr(excel_file, "seek"):
        excel_file.seek(0)
    raw_sheets = pd.read_excel(excel_file, sheet_name=sheet_names, header=None)
    return {name: parse_raw_count_sheet(raw, name) for name, raw in raw_sheets.items()}


def load_detected_sheets(excel_file: str | Path | BinaryIO) -> dict[str, pd.DataFrame]:
    return {name: parsed.data for name, parsed in load_detected_sheet_details(excel_file).items()}


def preview_detected_sheets(excel_file: str | Path | BinaryIO, preview_rows: int = 5) -> tuple[pd.DataFrame, dict[str, pd.DataFrame]]:
    if hasattr(excel_file, "seek"):
        excel_file.seek(0)
    detected = load_detected_sheet_details(excel_file)
    summary = [
        DetectedSheet(
            raw_sheet=name,
            raw_direction=parsed.debug.raw_direction,
            rows=len(parsed.raw),
            columns=len(parsed.raw.columns),
        ).__dict__
        for name, parsed in detected.items()
    ]
    previews = {name: parsed.raw.head(preview_rows) for name, parsed in detected.items()}
    return pd.DataFrame(summary), previews
