"""Mapping preset serialization helpers."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timezone
import json
from pathlib import Path
import re
from typing import Any

import pandas as pd

from .constants import MAPPING_COLUMNS
from .importer import extract_raw_direction
from .mapping import (
    OPTIONAL_MAPPING_METADATA_COLUMNS,
    clean_mapping,
    default_mapping_for_sheets,
    normalize_approach_movement_mapping,
    validate_mapping_scheme,
)
from .metadata import APP_VERSION
from .movement_scheme import MOVEMENT_SCHEME_V1, MOVEMENT_SCHEME_V2, normalize_movement_code_scheme


CURRENT_MAPPING_PRESET_SCHEMA_VERSION = 1
MAPPING_PRESET_TYPE = "tmc_mapping_preset"
MAPPING_PRESET_EXTENSION = ".mapping.json"
MAPPING_PRESET_MIME = "application/json"

MAPPING_PRESET_ROW_FIELDS = [
    "raw_sheet",
    "raw_direction",
    "source_stream",
    "raw_movement_label",
    "output_movement_code",
    "from_leg",
    "to_leg",
    "turn_type",
    "facility_type",
    "aggregation_method",
    "include_in_report",
    "include_in_peak",
    "note",
    "remark",
    "approach_direction",
    "movement_type",
]


class MappingPresetError(ValueError):
    """Raised when a mapping preset cannot be safely loaded."""


@dataclass(frozen=True)
class MappingPresetLoadResult:
    preset: dict[str, Any]
    warnings: tuple[str, ...] = ()


@dataclass(frozen=True)
class MappingPresetApplyResult:
    mapping: pd.DataFrame
    matched_sheet_count: int
    missing_detected_sheet_count: int
    extra_preset_row_count: int
    missing_detected_sheets: tuple[str, ...] = ()
    extra_preset_sheets: tuple[str, ...] = ()


@dataclass(frozen=True)
class MappingPresetMetadata:
    movement_code_scheme: str
    template_version: str = ""
    intersection_type: str = "four_leg"


def _utc_now_text() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")


def _json_safe(value: Any) -> Any:
    if isinstance(value, dict):
        return {str(key): _json_safe(item) for key, item in value.items()}
    if isinstance(value, (list, tuple)):
        return [_json_safe(item) for item in value]
    if isinstance(value, (datetime,)):
        return value.isoformat()
    if isinstance(value, Path):
        return str(value)
    if pd.isna(value) if not isinstance(value, (dict, list, tuple, set)) else False:
        return None
    if hasattr(value, "item"):
        try:
            return _json_safe(value.item())
        except (TypeError, ValueError):
            pass
    return value


def _preset_rows(mapping: Any, *, movement_code_scheme: str = MOVEMENT_SCHEME_V1) -> list[dict[str, Any]]:
    if mapping is None:
        return []
    frame = mapping if isinstance(mapping, pd.DataFrame) else pd.DataFrame(mapping)
    if frame.empty:
        return []

    note_values = {
        column: frame[column].fillna("").astype(str).tolist()
        for column in (*OPTIONAL_MAPPING_METADATA_COLUMNS, "note", "remark")
        if column in frame.columns
    }
    cleaned = clean_mapping(frame)
    rows: list[dict[str, Any]] = []
    for index, row in enumerate(cleaned.to_dict("records")):
        preset_row = {
            "raw_sheet": row.get("raw_sheet", ""),
            "raw_direction": row.get("raw_direction", ""),
            "source_stream": row.get("source_stream", ""),
            "raw_movement_label": row.get("raw_movement_label", ""),
            "output_movement_code": row.get("movement_code", ""),
            "from_leg": row.get("from_leg", ""),
            "to_leg": row.get("to_leg", ""),
            "turn_type": row.get("turn_type", ""),
            "facility_type": row.get("facility_type", ""),
            "aggregation_method": row.get("aggregation_method", ""),
            "include_in_report": row.get("include_in_report", True),
            "include_in_peak": row.get("include_in_peak", True),
        }
        for column, values in note_values.items():
            if index < len(values) and values[index].strip():
                preset_row[column] = values[index]
        rows.append(_json_safe(preset_row))
    return rows


def build_mapping_preset(
    mapping: Any,
    *,
    preset_name: str | None = None,
    notes: str | None = None,
    movement_code_scheme: str = MOVEMENT_SCHEME_V1,
    intersection_type: str = "four_leg",
    template_version: str | None = None,
    created_at: str | None = None,
    app_version: str | None = None,
) -> dict[str, Any]:
    """Build a JSON-serializable preset that contains mapping rows only."""

    scheme = normalize_movement_code_scheme(movement_code_scheme)
    return {
        "schema_version": CURRENT_MAPPING_PRESET_SCHEMA_VERSION,
        "preset_type": MAPPING_PRESET_TYPE,
        "preset_name": str(preset_name or "TMC Mapping Preset"),
        "created_at": created_at or _utc_now_text(),
        "app_version": app_version if app_version is not None else APP_VERSION,
        "movement_code_scheme": scheme,
        **({"template_version": str(template_version)} if template_version is not None else {}),
        "intersection_type": intersection_type,
        "notes": str(notes or ""),
        "mapping_rows": _preset_rows(mapping, movement_code_scheme=scheme),
    }


def serialize_mapping_preset(preset: dict[str, Any]) -> bytes:
    if not isinstance(preset, dict):
        raise MappingPresetError("Mapping preset must be a JSON object before serialization.")
    return json.dumps(_json_safe(preset), ensure_ascii=False, indent=2, sort_keys=True, allow_nan=False).encode("utf-8")


def load_mapping_preset(data: str | bytes | bytearray) -> MappingPresetLoadResult:
    try:
        text = bytes(data).decode("utf-8") if isinstance(data, (bytes, bytearray)) else str(data)
        raw = json.loads(text)
    except UnicodeDecodeError as exc:
        raise MappingPresetError("Mapping preset file must be UTF-8 encoded JSON.") from exc
    except json.JSONDecodeError as exc:
        raise MappingPresetError("Mapping preset file is not valid JSON.") from exc

    if not isinstance(raw, dict):
        raise MappingPresetError("Mapping preset JSON must be an object.")
    if raw.get("preset_type") != MAPPING_PRESET_TYPE:
        raise MappingPresetError("Mapping preset has an unsupported preset_type.")

    warnings: list[str] = []
    if raw.get("schema_version") != CURRENT_MAPPING_PRESET_SCHEMA_VERSION:
        warnings.append(
            f"Unsupported mapping preset schema_version {raw.get('schema_version')!r}; attempting to load safe fields only."
        )

    rows = raw.get("mapping_rows", [])
    if not isinstance(rows, list):
        rows = []
        warnings.append("Mapping preset mapping_rows was not a list and was ignored.")

    scheme = normalize_movement_code_scheme(raw.get("movement_code_scheme") or MOVEMENT_SCHEME_V1)
    normalized_rows = []
    for item in rows:
        if not isinstance(item, dict):
            continue
        row = {field: _json_safe(item.get(field, "")) for field in MAPPING_PRESET_ROW_FIELDS if field in item}
        if "output_movement_code" not in row and "movement_code" in item:
            row["output_movement_code"] = _json_safe(item.get("movement_code", ""))
        for key, value in item.items():
            if key not in row and key not in {"movement_code"}:
                row[str(key)] = _json_safe(value)
        normalized_rows.append(row)

    validation_issues = validate_mapping_preset_scheme(
        {"movement_code_scheme": scheme, "mapping_rows": normalized_rows}
    )
    if validation_issues:
        raise MappingPresetError("; ".join(validation_issues))

    known_top_level_fields = {
        "schema_version",
        "preset_type",
        "preset_name",
        "created_at",
        "app_version",
        "movement_code_scheme",
        "template_version",
        "intersection_type",
        "notes",
        "mapping_rows",
    }
    extra_metadata = {
        str(key): _json_safe(value)
        for key, value in raw.items()
        if key not in known_top_level_fields
    }

    preset = {
        "schema_version": raw.get("schema_version", CURRENT_MAPPING_PRESET_SCHEMA_VERSION),
        "preset_type": MAPPING_PRESET_TYPE,
        "preset_name": str(raw.get("preset_name") or ""),
        "created_at": str(raw.get("created_at") or ""),
        "app_version": str(raw.get("app_version") or ""),
        "movement_code_scheme": scheme,
        "template_version": str(raw.get("template_version") or ""),
        "intersection_type": str(raw.get("intersection_type") or "four_leg"),
        "notes": str(raw.get("notes") or ""),
        "mapping_rows": normalized_rows,
    }
    preset.update(extra_metadata)
    return MappingPresetLoadResult(preset=preset, warnings=tuple(warnings))


def detect_mapping_preset_scheme(preset: dict[str, Any] | MappingPresetLoadResult) -> str:
    preset_dict = preset.preset if isinstance(preset, MappingPresetLoadResult) else preset
    if not isinstance(preset_dict, dict):
        return MOVEMENT_SCHEME_V1
    return normalize_movement_code_scheme(preset_dict.get("movement_code_scheme") or MOVEMENT_SCHEME_V1)


def mapping_preset_metadata(preset: dict[str, Any] | MappingPresetLoadResult) -> MappingPresetMetadata:
    preset_dict = preset.preset if isinstance(preset, MappingPresetLoadResult) else preset
    source = preset_dict if isinstance(preset_dict, dict) else {}
    return MappingPresetMetadata(
        movement_code_scheme=normalize_movement_code_scheme(source.get("movement_code_scheme") or MOVEMENT_SCHEME_V1),
        template_version=str(source.get("template_version") or ""),
        intersection_type=str(source.get("intersection_type") or "four_leg"),
    )


def validate_mapping_preset_scheme(preset: dict[str, Any] | MappingPresetLoadResult) -> list[str]:
    preset_dict = preset.preset if isinstance(preset, MappingPresetLoadResult) else preset
    if not isinstance(preset_dict, dict):
        return []
    scheme = detect_mapping_preset_scheme(preset_dict)
    rows = preset_dict.get("mapping_rows", [])
    frame = _rows_to_mapping_frame(rows if isinstance(rows, list) else [], movement_code_scheme=scheme)
    return validate_mapping_scheme(frame, scheme)


def _rows_to_mapping_frame(
    rows: list[dict[str, Any]],
    *,
    movement_code_scheme: str = MOVEMENT_SCHEME_V1,
) -> pd.DataFrame:
    scheme = normalize_movement_code_scheme(movement_code_scheme)
    mapping_rows = []
    for row in rows:
        mapping_row = {
            "raw_sheet": row.get("raw_sheet", ""),
            "raw_direction": row.get("raw_direction", ""),
            "movement_code": row.get("output_movement_code", row.get("movement_code", "")),
            "source_stream": row.get("source_stream", ""),
            "raw_movement_label": row.get("raw_movement_label", ""),
            "from_leg": row.get("from_leg", ""),
            "to_leg": row.get("to_leg", ""),
            "turn_type": row.get("turn_type", ""),
            "facility_type": row.get("facility_type", ""),
            "include_in_peak": row.get("include_in_peak", True),
            "include_in_report": row.get("include_in_report", True),
            "aggregation_method": row.get("aggregation_method", ""),
        }
        for column in OPTIONAL_MAPPING_METADATA_COLUMNS:
            if column in row:
                mapping_row[column] = row.get(column, "")
        code = str(mapping_row["movement_code"] or "")
        if scheme != MOVEMENT_SCHEME_V2 and len(code) >= 2:
            mapping_row["from_leg"] = mapping_row["from_leg"] or code[:1]
            mapping_row["to_leg"] = mapping_row["to_leg"] or code[1:2]
        for column in ("note", "remark"):
            if str(row.get(column, "") or "").strip():
                mapping_row[column] = row[column]
        mapping_rows.append(mapping_row)
    if not mapping_rows:
        return pd.DataFrame(columns=MAPPING_COLUMNS)
    raw_frame = pd.DataFrame(mapping_rows)
    cleaned = normalize_approach_movement_mapping(raw_frame) if scheme == MOVEMENT_SCHEME_V2 else clean_mapping(raw_frame)
    for column in ("note", "remark"):
        if any(column in row for row in mapping_rows):
            cleaned[column] = [str(row.get(column, "") or "") for row in mapping_rows]
    return cleaned


def apply_mapping_preset_to_detected_sheets(
    preset: dict[str, Any] | MappingPresetLoadResult,
    detected_raw_sheets: list[str] | tuple[str, ...],
) -> MappingPresetApplyResult:
    """Apply preset rows to currently detected sheets while keeping unmatched rows out."""

    preset_dict = preset.preset if isinstance(preset, MappingPresetLoadResult) else preset
    rows = preset_dict.get("mapping_rows", []) if isinstance(preset_dict, dict) else []
    scheme = detect_mapping_preset_scheme(preset_dict) if isinstance(preset_dict, dict) else MOVEMENT_SCHEME_V1
    preset_frame = _rows_to_mapping_frame(rows if isinstance(rows, list) else [], movement_code_scheme=scheme)
    detected = [str(sheet) for sheet in detected_raw_sheets]
    detected_set = set(detected)

    active_rows: list[dict[str, Any]] = []
    matched_sheets: set[str] = set()
    extra_rows = 0
    extra_sheets: list[str] = []

    if not preset_frame.empty:
        for row in preset_frame.to_dict("records"):
            raw_sheet = str(row.get("raw_sheet") or "")
            if raw_sheet in detected_set:
                if not row.get("raw_direction"):
                    row["raw_direction"] = extract_raw_direction(raw_sheet) or ""
                active_rows.append(row)
                matched_sheets.add(raw_sheet)
            else:
                extra_rows += 1
                if raw_sheet and raw_sheet not in extra_sheets:
                    extra_sheets.append(raw_sheet)

    missing_sheets = [sheet for sheet in detected if sheet not in matched_sheets]
    for row in default_mapping_for_sheets(missing_sheets).to_dict("records"):
        row["remark"] = "Needs review: sheet was not found in the loaded mapping preset."
        active_rows.append(row)

    active = pd.DataFrame(active_rows)
    if active.empty:
        active = default_mapping_for_sheets(detected)
    else:
        cleaned = normalize_approach_movement_mapping(active) if scheme == MOVEMENT_SCHEME_V2 else clean_mapping(active)
        for column in ("note", "remark"):
            if column in active.columns:
                cleaned[column] = active[column].fillna("").astype(str).tolist()
        active = cleaned

    return MappingPresetApplyResult(
        mapping=active,
        matched_sheet_count=len(matched_sheets),
        missing_detected_sheet_count=len(missing_sheets),
        extra_preset_row_count=extra_rows,
        missing_detected_sheets=tuple(missing_sheets),
        extra_preset_sheets=tuple(extra_sheets),
    )


def safe_mapping_preset_filename(name: str | None, default: str = "tmc_mapping.mapping.json") -> str:
    cleaned = re.sub(r"[^A-Za-z0-9_-]+", "_", str(name or "").strip()).strip("_")
    if not cleaned:
        return default
    return f"{cleaned[:60]}{MAPPING_PRESET_EXTENSION}"
