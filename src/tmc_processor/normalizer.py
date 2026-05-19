"""Normalize raw vendor sheets into the AGENTS.md long schema."""

from __future__ import annotations

from typing import Any

import pandas as pd

from .constants import BASE_INTERVAL_MINUTES, NORMALIZED_COLUMNS, VEHICLE_CLASSES
from .mapping import clean_mapping
from .pcu import add_pcu
from .time_utils import add_minutes, parse_interval, parse_time


TIME_START_ALIASES = {"time_start", "start", "start_time", "\u0e40\u0e27\u0e25\u0e32\u0e40\u0e23\u0e34\u0e48\u0e21", "\u0e40\u0e23\u0e34\u0e48\u0e21"}
TIME_END_ALIASES = {"time_end", "end", "end_time", "\u0e40\u0e27\u0e25\u0e32\u0e2a\u0e34\u0e49\u0e19\u0e2a\u0e38\u0e14", "\u0e2a\u0e34\u0e49\u0e19\u0e2a\u0e38\u0e14"}
INTERVAL_ALIASES = {"time", "interval", "\u0e0a\u0e48\u0e27\u0e07\u0e40\u0e27\u0e25\u0e32", "\u0e40\u0e27\u0e25\u0e32"}
PARSED_RAW_COLUMNS = {"raw_sheet", "raw_direction", "time_start", "time_end", "vehicle_class", "count"}


def _normalize_column_name(column: Any) -> str:
    return str(column).strip()


def _find_column(columns: list[str], aliases: set[str]) -> str | None:
    lowered = {column.lower(): column for column in columns}
    for alias in aliases:
        if alias.lower() in lowered:
            return lowered[alias.lower()]
    for column in columns:
        lower = column.lower()
        if any(alias.lower() in lower for alias in aliases):
            return column
    return None


def _vehicle_columns(df: pd.DataFrame) -> list[str]:
    normalized = {_normalize_column_name(column): column for column in df.columns}
    return [normalized[vehicle] for vehicle in VEHICLE_CLASSES if vehicle in normalized]


def _extract_times(row: pd.Series, start_col: str | None, end_col: str | None, interval_col: str | None) -> tuple[Any, Any]:
    start = parse_time(row[start_col]) if start_col else None
    end = parse_time(row[end_col]) if end_col else None
    if interval_col and start is None:
        start, parsed_end = parse_interval(row[interval_col])
        end = end or parsed_end
    if start is not None and end is None:
        end = add_minutes(start, BASE_INTERVAL_MINUTES)
    return start, end


def _source_stream(map_row: pd.Series) -> str:
    text = str(map_row.get("source_stream", "") or "").strip()
    return text or "mainline"


def _raw_movement_label(map_row: pd.Series, raw_row: pd.Series) -> str:
    text = str(map_row.get("raw_movement_label", "") or "").strip()
    if text:
        return text
    return str(map_row.get("raw_direction", "") or raw_row.get("raw_direction", "") or "").strip()


def _append_normalized_row(rows: list[dict[str, Any]], setup: dict[str, Any], map_row: pd.Series, raw_sheet: str, raw_row: pd.Series) -> None:
    source_direction = map_row["raw_direction"] or raw_row.get("raw_direction", "")
    output_movement_code = str(map_row["movement_code"]).strip()
    rows.append(
        {
            "project": setup.get("project_name") or setup.get("project", ""),
            "tmc_id": setup.get("tmc_id", ""),
            "tmc_name": setup.get("tmc_name", ""),
            "survey_date": setup.get("survey_date", ""),
            "raw_sheet": raw_sheet,
            "raw_direction": source_direction,
            "source_direction": source_direction,
            "source_stream": _source_stream(map_row),
            "raw_movement_label": _raw_movement_label(map_row, raw_row),
            "movement_code": output_movement_code,
            "output_movement_code": output_movement_code,
            "from_leg": map_row["from_leg"],
            "to_leg": map_row["to_leg"],
            "turn_type": map_row["turn_type"],
            "facility_type": map_row["facility_type"],
            "include_in_peak": bool(map_row["include_in_peak"]),
            "include_in_report": bool(map_row["include_in_report"]),
            "time_start": raw_row["time_start"],
            "time_end": raw_row["time_end"],
            "vehicle_class": raw_row["vehicle_class"],
            "count": raw_row["count"],
        }
    )


def normalize(
    raw_sheets: dict[str, pd.DataFrame],
    mapping: pd.DataFrame,
    setup: dict[str, Any] | None = None,
    pce_factors: dict[str, float] | None = None,
) -> pd.DataFrame:
    setup = setup or {}
    mapping = clean_mapping(mapping)
    mapped = mapping[mapping["include_in_report"]].copy()
    rows: list[dict[str, Any]] = []

    for _, map_row in mapped.iterrows():
        raw_sheet = map_row["raw_sheet"]
        if raw_sheet not in raw_sheets:
            continue
        raw = raw_sheets[raw_sheet].copy()

        if PARSED_RAW_COLUMNS.issubset(set(raw.columns)):
            for _, raw_row in raw.iterrows():
                _append_normalized_row(rows, setup, map_row, raw_sheet, raw_row)
            continue

        raw.columns = [_normalize_column_name(column) for column in raw.columns]
        vehicle_cols = _vehicle_columns(raw)
        if not vehicle_cols:
            continue

        start_col = _find_column(list(raw.columns), TIME_START_ALIASES)
        end_col = _find_column(list(raw.columns), TIME_END_ALIASES)
        interval_col = _find_column(list(raw.columns), INTERVAL_ALIASES)

        for _, raw_row in raw.iterrows():
            time_start, time_end = _extract_times(raw_row, start_col, end_col, interval_col)
            if time_start is None:
                continue
            for vehicle_col in vehicle_cols:
                vehicle_class = _normalize_column_name(vehicle_col)
                parsed_row = pd.Series(
                    {
                        "raw_direction": map_row["raw_direction"],
                        "time_start": time_start,
                        "time_end": time_end,
                        "vehicle_class": vehicle_class,
                        "count": raw_row[vehicle_col],
                    }
                )
                _append_normalized_row(rows, setup, map_row, raw_sheet, parsed_row)

    normalized = pd.DataFrame(rows)
    if normalized.empty:
        return pd.DataFrame(columns=NORMALIZED_COLUMNS)
    normalized = add_pcu(normalized, pce_factors=pce_factors)
    return normalized[NORMALIZED_COLUMNS]
