"""Project session serialization helpers for reopenable TMC workflows."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import UTC, date, datetime, time
import json
from pathlib import Path
import re
from typing import Any, MutableMapping

import pandas as pd

from .constants import AM_WINDOW, DEFAULT_PEAK_MODE, MAPPING_COLUMNS, PM_WINDOW
from .mapping import clean_mapping
from .metadata import APP_VERSION, TEMPLATE_VERSION
from .pcu import normalize_pce_factors, validate_pce_factors


CURRENT_SCHEMA_VERSION = 1
SUPPORTED_SCHEMA_VERSIONS = {CURRENT_SCHEMA_VERSION}
PROJECT_SESSION_EXTENSION = ".tmcproj.json"
PROJECT_SESSION_MIME = "application/json"


METADATA_FIELDS = (
    "project_name",
    "tmc_id",
    "survey_point",
    "survey_date_text",
    "weather",
    "responsible_party",
    "survey_period",
    "tmc_title",
)

DIRECTION_FIELDS = (
    "north_label",
    "south_label",
    "east_label",
    "west_label",
    "north_road",
    "south_road",
    "east_road",
    "west_road",
    "caption_text",
    "show_u_turn",
)

PEAK_FIELDS = (
    "peak_mode",
    "am_peak_window_start",
    "am_peak_window_end",
    "pm_peak_window_start",
    "pm_peak_window_end",
    "am_peak_start",
    "am_peak_end",
    "pm_peak_start",
    "pm_peak_end",
    "peak_selection_source",
)

EXPORT_FIELDS = (
    "use_template_report_layout",
    "use_excel_com_native_charts",
    "template_version",
    "template_name",
    "template_path",
    "template_map_name",
    "template_map_path",
)


class ProjectSessionError(ValueError):
    """Raised when a project session file cannot be safely loaded."""


@dataclass(frozen=True)
class ProjectSessionLoadResult:
    session: dict[str, Any]
    warnings: tuple[str, ...] = ()


def _utc_now_text() -> str:
    return datetime.now(UTC).replace(microsecond=0).isoformat().replace("+00:00", "Z")


def _time_text(value: Any) -> str:
    if isinstance(value, time):
        return value.strftime("%H:%M")
    text = str(value or "").strip()
    return text[:5] if re.fullmatch(r"\d{2}:\d{2}(:\d{2})?", text) else text


def _parse_time(value: Any) -> time | None:
    if isinstance(value, time):
        return value
    text = str(value or "").strip()
    if not text:
        return None
    try:
        hour, minute = text[:5].split(":")
        return time(int(hour), int(minute))
    except (TypeError, ValueError):
        return None


def _json_safe(value: Any) -> Any:
    if isinstance(value, dict):
        return {str(key): _json_safe(item) for key, item in value.items()}
    if isinstance(value, (list, tuple)):
        return [_json_safe(item) for item in value]
    if isinstance(value, (datetime, date, time)):
        return value.isoformat()
    if isinstance(value, Path):
        return str(value)
    if pd.isna(value) if not isinstance(value, (list, tuple, dict, set)) else False:
        return None
    if hasattr(value, "item"):
        try:
            return _json_safe(value.item())
        except (TypeError, ValueError):
            pass
    return value


def _copy_known_fields(values: dict[str, Any] | None, fields: tuple[str, ...]) -> dict[str, Any]:
    source = values if isinstance(values, dict) else {}
    return {field: _json_safe(source[field]) for field in fields if field in source}


def _mapping_rows(mapping: Any) -> list[dict[str, Any]]:
    if mapping is None:
        return []
    if isinstance(mapping, pd.DataFrame):
        frame = mapping
    elif isinstance(mapping, list):
        frame = pd.DataFrame(mapping)
    else:
        return []
    if frame.empty:
        return []
    cleaned = clean_mapping(frame)
    return [_json_safe(row) for row in cleaned.to_dict("records")]


def build_project_session(
    *,
    metadata: dict[str, Any] | None = None,
    directions: dict[str, Any] | None = None,
    mapping: Any = None,
    detected_sheet_names: list[str] | tuple[str, ...] | None = None,
    peak_settings: dict[str, Any] | None = None,
    export_settings: dict[str, Any] | None = None,
    pce_factors: dict[str, Any] | None = None,
    source_file_name: str | None = None,
    source_file_size: int | None = None,
    source_file_modified_time: str | None = None,
    created_at: str | None = None,
    updated_at: str | None = None,
) -> dict[str, Any]:
    """Build a JSON-serializable project session without embedding workbook bytes."""

    now = _utc_now_text()
    peak_values = _copy_known_fields(peak_settings, PEAK_FIELDS)
    for key in (
        "am_peak_window_start",
        "am_peak_window_end",
        "pm_peak_window_start",
        "pm_peak_window_end",
        "am_peak_start",
        "am_peak_end",
        "pm_peak_start",
        "pm_peak_end",
    ):
        if key in peak_values:
            peak_values[key] = _time_text(peak_values[key])

    return {
        "schema_version": CURRENT_SCHEMA_VERSION,
        "app_version": APP_VERSION,
        "template_version": TEMPLATE_VERSION,
        "created_at": created_at or now,
        "updated_at": updated_at or now,
        "source_file_name": source_file_name or "",
        "source_file_size": source_file_size,
        "source_file_modified_time": source_file_modified_time or "",
        "metadata": _copy_known_fields(metadata, METADATA_FIELDS),
        "directions": _copy_known_fields(directions, DIRECTION_FIELDS),
        "mapping": {
            "detected_sheet_names": [str(name) for name in (detected_sheet_names or [])],
            "rows": _mapping_rows(mapping),
        },
        "pce_factors": _json_safe(normalize_pce_factors(pce_factors)),
        "peaks": peak_values,
        "export": _copy_known_fields(export_settings, EXPORT_FIELDS),
    }


def normalize_project_session(raw_session: dict[str, Any]) -> dict[str, Any]:
    """Return only known, safe fields from a loaded project session."""

    mapping_source = raw_session.get("mapping") if isinstance(raw_session.get("mapping"), dict) else {}
    rows = mapping_source.get("rows", [])
    detected = mapping_source.get("detected_sheet_names", [])
    if not isinstance(rows, list):
        rows = []
    if not isinstance(detected, list):
        detected = []

    return {
        "schema_version": raw_session.get("schema_version", CURRENT_SCHEMA_VERSION),
        "app_version": str(raw_session.get("app_version") or ""),
        "template_version": str(raw_session.get("template_version") or ""),
        "created_at": str(raw_session.get("created_at") or ""),
        "updated_at": str(raw_session.get("updated_at") or ""),
        "source_file_name": str(raw_session.get("source_file_name") or ""),
        "source_file_size": raw_session.get("source_file_size"),
        "source_file_modified_time": str(raw_session.get("source_file_modified_time") or ""),
        "metadata": _copy_known_fields(raw_session.get("metadata"), METADATA_FIELDS),
        "directions": _copy_known_fields(raw_session.get("directions"), DIRECTION_FIELDS),
        "mapping": {
            "detected_sheet_names": [str(name) for name in detected],
            "rows": _mapping_rows(rows),
        },
        "pce_factors": _json_safe(normalize_pce_factors(raw_session.get("pce_factors"))),
        "peaks": _copy_known_fields(raw_session.get("peaks"), PEAK_FIELDS),
        "export": _copy_known_fields(raw_session.get("export"), EXPORT_FIELDS),
    }


def session_to_json(session: dict[str, Any]) -> str:
    """Serialize a project session as deterministic, pretty JSON text."""

    if not isinstance(session, dict):
        raise ProjectSessionError("Project session must be a JSON object before serialization.")
    return json.dumps(_json_safe(session), ensure_ascii=False, indent=2, sort_keys=True, allow_nan=False)


def session_to_json_bytes(session: dict[str, Any]) -> bytes:
    return session_to_json(session).encode("utf-8")


def session_from_json(data: str | bytes | bytearray) -> ProjectSessionLoadResult:
    """Parse, validate, and sanitize a project session JSON payload."""

    try:
        text = bytes(data).decode("utf-8") if isinstance(data, (bytes, bytearray)) else str(data)
        raw = json.loads(text)
    except UnicodeDecodeError as exc:
        raise ProjectSessionError("Project session file must be UTF-8 encoded JSON.") from exc
    except json.JSONDecodeError as exc:
        raise ProjectSessionError("Project session file is not valid JSON.") from exc

    if not isinstance(raw, dict):
        raise ProjectSessionError("Project session JSON must be an object.")

    warnings: list[str] = []
    schema_version = raw.get("schema_version", CURRENT_SCHEMA_VERSION)
    if schema_version not in SUPPORTED_SCHEMA_VERSIONS:
        warnings.append(
            f"Unsupported project session schema_version {schema_version!r}; attempting to load safe fields only."
        )
    pce_validation = validate_pce_factors(raw.get("pce_factors"))
    warnings.extend(pce_validation.warnings)
    return ProjectSessionLoadResult(session=normalize_project_session(raw), warnings=tuple(warnings))


def safe_project_session_filename(name: str | None, default: str = "tmc_session.tmcproj.json") -> str:
    cleaned = re.sub(r"[^A-Za-z0-9_-]+", "_", str(name or "").strip()).strip("_")
    if not cleaned:
        return default
    return f"{cleaned[:60]}_session{PROJECT_SESSION_EXTENSION}"


def source_file_mismatch_warning(session: dict[str, Any], uploaded_file_name: str | None) -> str | None:
    session_name = str(session.get("source_file_name") or "").strip()
    uploaded_name = str(uploaded_file_name or "").strip()
    if session_name and uploaded_name and session_name != uploaded_name:
        return "This session was created for a different source file. Please verify before processing."
    return None


def apply_session_to_state(session: dict[str, Any], state: MutableMapping[str, Any]) -> list[str]:
    """Apply loaded session settings to a Streamlit-like state mapping."""

    updates: dict[str, Any] = {}
    for field, key in {
        "project_name": "project_name_input",
        "tmc_id": "tmc_id_input",
        "survey_point": "survey_point_input",
        "survey_date_text": "survey_date_text_input",
        "weather": "weather_input",
        "responsible_party": "responsible_party_input",
        "survey_period": "survey_period_input",
        "tmc_title": "tmc_title_input",
    }.items():
        if field in session.get("metadata", {}):
            updates[key] = session["metadata"][field]

    for field in DIRECTION_FIELDS:
        if field in session.get("directions", {}):
            key = "show_u_turn_checkbox" if field == "show_u_turn" else f"{field}_input"
            updates[key] = session["directions"][field]

    peaks = session.get("peaks", {})
    if "peak_mode" in peaks:
        updates["peak_mode_select"] = peaks["peak_mode"] or DEFAULT_PEAK_MODE
    for field in ("am_peak_window_start", "am_peak_window_end", "pm_peak_window_start", "pm_peak_window_end"):
        if field in peaks:
            parsed = _parse_time(peaks[field])
            if parsed is not None:
                updates[f"{field}_input"] = parsed
    confirmed = {
        key: peaks.get(key, "")
        for key in ("am_peak_start", "am_peak_end", "pm_peak_start", "pm_peak_end", "peak_selection_source")
        if key in peaks
    }
    if confirmed:
        updates["tmc_loaded_confirmed_peaks"] = confirmed

    export = session.get("export", {})
    for field, key in {
        "use_template_report_layout": "use_template_report_layout_checkbox",
        "use_excel_com_native_charts": "use_excel_com_native_charts_checkbox",
        "template_name": "template_name",
        "template_path": "template_path",
        "template_map_name": "template_map_name",
        "template_map_path": "template_map_path",
    }.items():
        if field in export:
            updates[key] = export[field]

    rows = session.get("mapping", {}).get("rows", [])
    if rows:
        updates["mapping_table"] = rows
        updates["mapping_editor_version"] = int(state.get("mapping_editor_version", 0) or 0) + 1
        updates["tmc_mapping_table_from_session"] = True
    detected = session.get("mapping", {}).get("detected_sheet_names", [])
    if detected:
        updates["tmc_session_detected_sheet_names"] = detected

    pce_factors = session.get("pce_factors")
    if isinstance(pce_factors, dict):
        normalized_pce = normalize_pce_factors(pce_factors)
        updates["pce_factors_table"] = [
            {"vehicle_class": vehicle_class, "pce_factor": factor}
            for vehicle_class, factor in normalized_pce.items()
        ]
        updates["pce_editor_version"] = int(state.get("pce_editor_version", 0) or 0) + 1

    changed: list[str] = []
    for key, value in updates.items():
        if state.get(key) != value:
            changed.append(key)
        state[key] = value
    return changed


def default_peak_window_state() -> dict[str, time]:
    return {
        "am_peak_window_start_input": _parse_time(AM_WINDOW[0]) or time(7, 0),
        "am_peak_window_end_input": _parse_time(AM_WINDOW[1]) or time(12, 0),
        "pm_peak_window_start_input": _parse_time(PM_WINDOW[0]) or time(15, 0),
        "pm_peak_window_end_input": _parse_time(PM_WINDOW[1]) or time(19, 0),
    }
