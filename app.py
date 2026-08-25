"""Streamlit UI for the TMC Processor MVP."""

from __future__ import annotations

from datetime import time
import hashlib
from html import escape
from io import BytesIO
from pathlib import Path
import sys
from typing import MutableMapping
import warnings

import altair as alt
import pandas as pd
import streamlit as st

from tmc_processor.batch import (
    BATCH_EXCEL_TEMPLATE_EXPORT_MODE,
    BATCH_PACKAGE_MIME,
    BATCH_SAFE_PNG_EXPORT_MODE,
    BATCH_V2_TEMPLATE_MODE_UNSUPPORTED_TH,
    BatchItem,
    analyze_batch_files,
    batch_analysis_qc_rows,
    batch_folder_name,
    batch_change_invalidates,
    batch_file_metadata_defaults,
    batch_inputs_ready,
    batch_package_filename,
    batch_processing_block_reason,
    batch_qc_frame,
    batch_selected_file_preview,
    batch_zip_generation_block_reason,
    batch_zip_contents_preview,
    generate_batch_zip_from_reviewed_peaks,
    reviewed_peak_values_complete,
    safe_output_stem,
    unique_safe_output_stems,
)
from tmc_processor.charts import report_chart_pngs
from tmc_processor.constants import (
    AM_WINDOW,
    AGGREGATION_METHOD_OPTIONS,
    DEFAULT_PEAK_MODE,
    FACILITY_TYPE_OPTIONS,
    LEG_OPTIONS,
    MOVEMENT_CODE_OPTIONS,
    PEAK_MODE_OPTIONS,
    PM_WINDOW,
    SOURCE_STREAM_OPTIONS,
    TURN_TYPE_OPTIONS,
)
from tmc_processor.diagram import DiagramConfig, build_v2_movement_diagram_data, generate_four_leg_tmc_diagram
from tmc_processor.downloads import EXCEL_MIME, PNG_MIME, download_buffer, safe_workbook_filename
from tmc_processor.excel_com_export import ExcelComStatus, probe_excel_com
from tmc_processor.export_package import (
    PACKAGE_MIME,
    build_export_summary_text,
    create_export_package_zip,
    create_v2_generated_export_package_zip,
    safe_package_filename,
)
from tmc_processor.importer import detect_raw_direction_sheet_names, load_detected_sheet_details, preview_detected_sheets
from tmc_processor.mapping import (
    apply_saved_mapping_to_sheets,
    default_mapping_for_sheets,
    mapping_to_excel_bytes,
    mapping_control_warnings,
    mapping_is_process_compatible,
    mapping_processing_block_reason,
    normalize_mapping_for_scheme,
    normalize_approach_movement_mapping,
    movement_aggregation_messages,
    read_mapping_excel_with_metadata,
    selectbox_options_with_existing_values,
    validate_mapping_scheme,
    validate_mapping_for_processing_by_scheme,
)
from tmc_processor.mapping_preset import (
    MAPPING_PRESET_MIME,
    MappingPresetError,
    apply_mapping_preset_to_detected_sheets,
    build_mapping_preset,
    detect_mapping_preset_scheme,
    load_mapping_preset,
    safe_mapping_preset_filename,
    serialize_mapping_preset,
)
from tmc_processor.metadata import (
    APP_VERSION,
    DEFAULT_CAPTION_TEXT,
    DEFAULT_RESPONSIBLE_PARTY,
    DEFAULT_SURVEY_PERIOD,
    DEFAULT_WEATHER,
    SetupMetadata,
    TEMPLATE_VERSION,
    generated_timestamp_text,
)
from tmc_processor.pcu import (
    get_default_pce_factors,
    normalize_pce_factors,
    pce_factor_traceability_frame,
    pce_factors_equal,
    validate_pce_factors,
)
from tmc_processor.peaks import PEAK_SELECTION_AUTO, PEAK_SELECTION_USER_CONFIRMED
from tmc_processor.pipeline import process_tmc, process_tmc_dry_run_v2
from tmc_processor.movement_scheme import (
    APPROACH_MOVEMENT_CODES,
    MOVEMENT_SCHEME_V1,
    MOVEMENT_SCHEME_V2,
    MOVEMENT_SCHEMES,
    movement_scheme_description,
    movement_scheme_display_label,
    movement_scheme_status_label,
    normalize_movement_code_scheme,
)
from tmc_processor.report_template import DEFAULT_TEMPLATE_MAP_PATH, DEFAULT_TEMPLATE_PATH, load_template_map
from tmc_processor.exporter import export_v2_generated_workbook, export_v2_template_workbook_com
from tmc_processor.session import (
    PROJECT_SESSION_MIME,
    ProjectSessionError,
    apply_session_to_state,
    build_project_session,
    default_peak_window_state,
    safe_project_session_filename,
    session_from_json,
    session_to_json_bytes,
    source_file_mismatch_warning,
)
from tmc_processor.summaries import hourly_movement_pcu, movement_aggregation_audit, vehicle_composition_report
from tmc_processor.template_audit import audit_template_formulas
from tmc_processor.time_utils import (
    hourly_interval_label_parts,
    hourly_interval_options as base_hourly_interval_options,
    hourly_interval_rows,
)
from tmc_processor.workflow_state import (
    WorkflowReadiness,
    WorkflowRevisions,
    WorkflowState,
    WorkflowTransition,
    analysis_config_fingerprint,
    analysis_result_fingerprint,
    export_config_fingerprint,
    mapping_fingerprint,
    review_decision_fingerprint,
    readiness_after_transition,
    semantic_fingerprint,
    source_fingerprint,
    transition_workflow,
)


EXCEL_TEMPLATE_EXPORT_MODE = "Excel Template Mode — แนะนำ"
SAFE_PNG_EXPORT_MODE = "Safe PNG Export Mode — โหมดสำรอง"
V2_EXCEL_TEMPLATE_MODE_BLOCK_MESSAGE = (
    "Excel Template Mode สำหรับ approach_movement ต้องใช้ Excel COM เพื่อรักษากราฟและ diagram จาก template "
    "ขณะนี้ไม่สามารถใช้โหมดนี้ได้ กรุณาใช้ Safe PNG Export Mode หรือเปิดใช้งาน Excel COM"
)
BATCH_EXCEL_TEMPLATE_EXPORT_LABEL = EXCEL_TEMPLATE_EXPORT_MODE
BATCH_SAFE_PNG_EXPORT_LABEL = "Safe PNG Export Mode — โหมดสำรอง"
WORKFLOW_TAB_LABELS = ["ตั้งค่า", "กำหนดทิศทาง", "ตรวจ Peak", "ส่งออก", "ตรวจสอบข้อมูล"]
DEFAULT_WORKFLOW_TAB = "ตั้งค่า"
SETUP_STATE_KEY = "tmc_setup_state"
SETUP_FIELD_WIDGET_KEYS = {
    "project_name": "project_name_input",
    "tmc_id": "tmc_id_input",
    "tmc_title": "tmc_title_input",
    "survey_point": "survey_point_input",
    "survey_date_text": "survey_date_text_input",
    "weather": "weather_input",
    "responsible_party": "responsible_party_input",
    "survey_period": "survey_period_input",
    "north_label": "north_label_input",
    "south_label": "south_label_input",
    "east_label": "east_label_input",
    "west_label": "west_label_input",
    "north_road": "north_road_input",
    "south_road": "south_road_input",
    "east_road": "east_road_input",
    "west_road": "west_road_input",
    "caption_text": "caption_text_input",
    "peak_mode": "peak_mode_select",
    "am_peak_window_start": "am_peak_window_start_input",
    "am_peak_window_end": "am_peak_window_end_input",
    "pm_peak_window_start": "pm_peak_window_start_input",
    "pm_peak_window_end": "pm_peak_window_end_input",
    "show_u_turn": "show_u_turn_checkbox",
}
SETUP_TIME_FIELDS = {
    "am_peak_window_start",
    "am_peak_window_end",
    "pm_peak_window_start",
    "pm_peak_window_end",
}

WORKFLOW_STATE_KEY = "tmc_workflow_state"
WORKFLOW_SINGLE_MODE = "single"
WORKFLOW_BATCH_MODE = "batch"
WORKFLOW_EXPORT_METADATA_FIELDS = (
    "project_name",
    "tmc_id",
    "tmc_title",
    "survey_point",
    "survey_date_text",
    "weather",
    "responsible_party",
    "survey_period",
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

CHART_PRIMARY_COLOR = "#0E4A2A"
CHART_PM_COLOR = "#B57A22"
CHART_GRID_COLOR = "#E6E1D8"
CHART_TEXT_COLOR = "#151713"


def _default_text_from_filename(filename: str | None) -> str:
    if not filename:
        return ""
    return Path(filename).stem.replace("_", " ").replace("-", " ").strip()


def _time_from_text(value: str) -> time:
    hour, minute = value.split(":")
    return time(int(hour), int(minute))


def _time_text(value: time) -> str:
    return value.strftime("%H:%M")


def _coerce_display_number(value: object) -> float | None:
    numeric = pd.to_numeric(pd.Series([value]), errors="coerce").iloc[0]
    return float(numeric) if pd.notna(numeric) else None


def format_count(value: object) -> str:
    numeric = _coerce_display_number(value)
    if numeric is None:
        return "-"
    return f"{numeric:,.0f}"


def format_pcu(value: object) -> str:
    numeric = _coerce_display_number(value)
    if numeric is None:
        return "-"
    return f"{numeric:,.0f}" if abs(numeric - round(numeric)) < 0.005 else f"{numeric:,.1f}"


def format_percent(value: object) -> str:
    numeric = _coerce_display_number(value)
    if numeric is None:
        return "-"
    percent = numeric * 100 if abs(numeric) <= 1 else numeric
    return f"{percent:,.1f}%"


def format_phf(value: object) -> str:
    numeric = _coerce_display_number(value)
    if numeric is None:
        return "-"
    return f"{numeric:.2f}"


def format_time_range(value: object) -> str:
    text = str(value or "").strip()
    if not text:
        return "-"
    parts = hourly_interval_label_parts(text)
    if parts is None:
        return text.replace(".", ":")
    return f"{parts[0]}–{parts[1]}"


def format_vehicle_unit(value: object, unit: str = "คัน") -> str:
    return f"{format_count(value)} {unit}"


def _setup_from_inputs(
    *,
    project_name: str,
    tmc_id: str,
    tmc_title: str,
    survey_point: str,
    survey_date_text: str,
    weather: str,
    responsible_party: str,
    survey_period: str,
    north_label: str,
    south_label: str,
    east_label: str,
    west_label: str,
    north_road: str,
    south_road: str,
    east_road: str,
    west_road: str,
    caption_text: str,
    uploaded_filename: str,
    peak_mode: str,
    am_peak_window_start: time,
    am_peak_window_end: time,
    pm_peak_window_start: time,
    pm_peak_window_end: time,
    show_u_turn: bool,
) -> dict[str, object]:
    metadata = SetupMetadata(
        project_name=project_name,
        survey_point=survey_point,
        survey_date_text=survey_date_text,
        weather=weather,
        responsible_party=responsible_party,
        survey_period=survey_period,
        tmc_title=tmc_title,
        north_label=north_label,
        south_label=south_label,
        east_label=east_label,
        west_label=west_label,
        north_road=north_road,
        south_road=south_road,
        east_road=east_road,
        west_road=west_road,
        caption_text=caption_text,
    )
    return metadata.to_setup_dict(
        extra={
            "tmc_id": tmc_id,
            "source_file": uploaded_filename,
            "peak_mode": peak_mode,
            "am_peak_window_start": _time_text(am_peak_window_start),
            "am_peak_window_end": _time_text(am_peak_window_end),
            "pm_peak_window_start": _time_text(pm_peak_window_start),
            "pm_peak_window_end": _time_text(pm_peak_window_end),
            "show_u_turn": show_u_turn,
        }
    )


def _setup_state_defaults(uploaded_filename_default: str = "") -> dict[str, object]:
    peak_defaults = default_peak_window_state()
    return {
        "project_name": "",
        "tmc_id": "",
        "tmc_title": uploaded_filename_default,
        "survey_point": uploaded_filename_default,
        "survey_date_text": "",
        "weather": DEFAULT_WEATHER,
        "responsible_party": DEFAULT_RESPONSIBLE_PARTY,
        "survey_period": DEFAULT_SURVEY_PERIOD,
        "north_label": "",
        "south_label": "",
        "east_label": "",
        "west_label": "",
        "north_road": "",
        "south_road": "",
        "east_road": "",
        "west_road": "",
        "caption_text": DEFAULT_CAPTION_TEXT,
        "peak_mode": DEFAULT_PEAK_MODE,
        "am_peak_window_start": peak_defaults["am_peak_window_start_input"],
        "am_peak_window_end": peak_defaults["am_peak_window_end_input"],
        "pm_peak_window_start": peak_defaults["pm_peak_window_start_input"],
        "pm_peak_window_end": peak_defaults["pm_peak_window_end_input"],
        "show_u_turn": True,
    }


def _setup_state_mapping(state: MutableMapping[str, object] | None = None) -> MutableMapping[str, object]:
    return st.session_state if state is None else state


def _coerce_setup_time(value: object, default: time) -> time:
    if isinstance(value, time):
        return value
    text = str(value or "").strip()
    if text:
        try:
            hour, minute = text[:5].split(":")
            return time(int(hour), int(minute))
        except (TypeError, ValueError):
            pass
    return default


def _coerce_setup_value(field: str, value: object) -> object:
    defaults = _setup_state_defaults()
    if field == "show_u_turn":
        return bool(value)
    if field == "peak_mode":
        return value if value in PEAK_MODE_OPTIONS else DEFAULT_PEAK_MODE
    if field in SETUP_TIME_FIELDS:
        return _coerce_setup_time(value, defaults[field])
    if value is None:
        return ""
    return value


def _normalized_setup_state(values: dict[str, object] | None, uploaded_filename_default: str = "") -> dict[str, object]:
    defaults = _setup_state_defaults(uploaded_filename_default)
    normalized = dict(defaults)
    for field in SETUP_FIELD_WIDGET_KEYS:
        if isinstance(values, dict) and field in values:
            normalized[field] = _coerce_setup_value(field, values[field])
    return normalized


def update_setup_state_from_widgets(state: MutableMapping[str, object] | None = None) -> list[str]:
    state = _setup_state_mapping(state)
    setup_state = _normalized_setup_state(state.get(SETUP_STATE_KEY) if isinstance(state.get(SETUP_STATE_KEY), dict) else None)
    changed: list[str] = []
    for field, widget_key in SETUP_FIELD_WIDGET_KEYS.items():
        if widget_key not in state:
            continue
        value = _coerce_setup_value(field, state[widget_key])
        if setup_state.get(field) != value:
            changed.append(field)
        setup_state[field] = value
    state[SETUP_STATE_KEY] = setup_state
    return changed


def _hydrate_setup_widgets_from_state(state: MutableMapping[str, object] | None = None, *, force: bool = False) -> None:
    state = _setup_state_mapping(state)
    setup_state = _normalized_setup_state(state.get(SETUP_STATE_KEY) if isinstance(state.get(SETUP_STATE_KEY), dict) else None)
    state[SETUP_STATE_KEY] = setup_state
    for field, widget_key in SETUP_FIELD_WIDGET_KEYS.items():
        if force or widget_key not in state:
            state[widget_key] = setup_state[field]


def initialize_setup_state_once(
    uploaded_filename_default: str = "",
    state: MutableMapping[str, object] | None = None,
) -> dict[str, object]:
    state = _setup_state_mapping(state)
    existing = state.get(SETUP_STATE_KEY)
    if isinstance(existing, dict):
        state[SETUP_STATE_KEY] = _normalized_setup_state(existing, uploaded_filename_default)
    else:
        state[SETUP_STATE_KEY] = _setup_state_defaults(uploaded_filename_default)
    update_setup_state_from_widgets(state)
    _hydrate_setup_widgets_from_state(state)
    return dict(state[SETUP_STATE_KEY])


def apply_uploaded_filename_defaults_to_setup_state(
    uploaded_filename_default: str,
    state: MutableMapping[str, object] | None = None,
) -> list[str]:
    state = _setup_state_mapping(state)
    if not uploaded_filename_default:
        return []
    initialize_setup_state_once(uploaded_filename_default, state)
    setup_state = dict(state[SETUP_STATE_KEY])
    changed: list[str] = []
    for field in ("tmc_title", "survey_point"):
        widget_key = SETUP_FIELD_WIDGET_KEYS[field]
        if not str(setup_state.get(field, "") or "").strip():
            setup_state[field] = uploaded_filename_default
            changed.append(field)
        if not str(state.get(widget_key, "") or "").strip():
            state[widget_key] = uploaded_filename_default
    state[SETUP_STATE_KEY] = setup_state
    return changed


def get_current_setup_from_state(
    uploaded_filename: str = "",
    state: MutableMapping[str, object] | None = None,
) -> dict[str, object]:
    state = _setup_state_mapping(state)
    if SETUP_STATE_KEY not in state:
        initialize_setup_state_once(state=state)
    update_setup_state_from_widgets(state)
    values = _normalized_setup_state(state.get(SETUP_STATE_KEY) if isinstance(state.get(SETUP_STATE_KEY), dict) else None)
    return _setup_from_inputs(
        project_name=str(values["project_name"] or ""),
        tmc_id=str(values["tmc_id"] or ""),
        tmc_title=str(values["tmc_title"] or ""),
        survey_point=str(values["survey_point"] or ""),
        survey_date_text=str(values["survey_date_text"] or ""),
        weather=str(values["weather"] or ""),
        responsible_party=str(values["responsible_party"] or ""),
        survey_period=str(values["survey_period"] or ""),
        north_label=str(values["north_label"] or ""),
        south_label=str(values["south_label"] or ""),
        east_label=str(values["east_label"] or ""),
        west_label=str(values["west_label"] or ""),
        north_road=str(values["north_road"] or ""),
        south_road=str(values["south_road"] or ""),
        east_road=str(values["east_road"] or ""),
        west_road=str(values["west_road"] or ""),
        caption_text=str(values["caption_text"] or ""),
        uploaded_filename=uploaded_filename,
        peak_mode=str(values["peak_mode"] or DEFAULT_PEAK_MODE),
        am_peak_window_start=_coerce_setup_time(values["am_peak_window_start"], _time_from_text(AM_WINDOW[0])),
        am_peak_window_end=_coerce_setup_time(values["am_peak_window_end"], _time_from_text(AM_WINDOW[1])),
        pm_peak_window_start=_coerce_setup_time(values["pm_peak_window_start"], _time_from_text(PM_WINDOW[0])),
        pm_peak_window_end=_coerce_setup_time(values["pm_peak_window_end"], _time_from_text(PM_WINDOW[1])),
        show_u_turn=bool(values["show_u_turn"]),
    )


def build_setup_for_processing(
    uploaded_filename: str = "",
    state: MutableMapping[str, object] | None = None,
) -> dict[str, object]:
    return get_current_setup_from_state(uploaded_filename, state)


def _peak_row(peaks: pd.DataFrame, period: str) -> pd.Series | None:
    if peaks.empty or "period" not in peaks:
        return None
    matches = peaks[peaks["period"].astype(str).str.upper() == period]
    return matches.iloc[0] if not matches.empty else None


def _peak_period_text(peaks: pd.DataFrame, period: str) -> tuple[str, str, str]:
    row = _peak_row(peaks, period)
    if row is None:
        return "", "", ""
    start = row.get("peak_start")
    end = row.get("peak_end")
    start_text = start.strftime("%H:%M") if isinstance(start, time) else str(start)[:5]
    end_text = end.strftime("%H:%M") if isinstance(end, time) else str(end)[:5]
    pcu = pd.to_numeric(pd.Series([row.get("hourly_pcu")]), errors="coerce").fillna(0).iloc[0]
    return start_text, end_text, format_pcu(pcu)


def _hourly_interval_options(hourly_movement: pd.DataFrame, peaks: pd.DataFrame) -> list[tuple[str, str, str]]:
    options = base_hourly_interval_options(hourly_movement)

    existing = {(start, end) for _, start, end in options}
    for period in ["AM", "PM"]:
        start, end, _ = _peak_period_text(peaks, period)
        if start and end and (start, end) not in existing and hourly_interval_label_parts(f"{start}-{end}") is not None:
            options.append((f"{start}-{end}", start, end))
            existing.add((start, end))
    return options


def _selected_interval(options: list[tuple[str, str, str]], label: str) -> tuple[str, str]:
    for option_label, start, end in options:
        if option_label == label:
            return start, end
    return "", ""


def _render_download_button(label: str, data: bytes | None, file_name: str, mime: str) -> None:
    if not data:
        st.error(f"{label} ไม่พร้อมใช้งาน เนื่องจากไฟล์ที่สร้างไม่มีข้อมูล")
        return
    st.download_button(
        label,
        data=download_buffer(data),
        file_name=file_name,
        mime=mime,
    )


def set_active_tab(tab_name: str) -> str:
    active = tab_name if tab_name in WORKFLOW_TAB_LABELS else DEFAULT_WORKFLOW_TAB
    st.session_state["active_workflow_tab"] = active
    return active


def get_active_tab() -> str:
    return set_active_tab(str(st.session_state.get("active_workflow_tab") or DEFAULT_WORKFLOW_TAB))


def workflow_tab_choices() -> list[str]:
    return list(WORKFLOW_TAB_LABELS)


def _workflow_tab_button_key(index: int) -> str:
    return f"workflow_tab_{index}"


def render_workflow_tab_nav(active_tab: str, tabs: list[str]) -> str:
    available_tabs = list(tabs)
    if not available_tabs:
        return set_active_tab(DEFAULT_WORKFLOW_TAB)

    active = active_tab if active_tab in available_tabs else DEFAULT_WORKFLOW_TAB
    if active not in available_tabs:
        active = available_tabs[0]

    tab_columns = st.columns(len(available_tabs), gap="small")
    for index, tab_name in enumerate(available_tabs):
        with tab_columns[index]:
            clicked = st.button(
                tab_name,
                key=_workflow_tab_button_key(index),
                type="primary" if tab_name == active else "secondary",
                use_container_width=True,
            )
        if clicked and tab_name != active:
            set_active_tab(tab_name)
            st.rerun()

    return active


def render_workflow_navigation() -> str:
    active_tab = get_active_tab()
    selected_tab = render_workflow_tab_nav(active_tab, workflow_tab_choices())
    return set_active_tab(selected_tab)


def apply_single_export_mode_change(selected_mode: str, previous_mode: str | None) -> bool:
    set_active_tab("ส่งออก")
    if selected_mode == previous_mode:
        return False
    st.session_state["report_export_mode"] = selected_mode
    _clear_single_export()
    return True


def _clear_single_export() -> bool:
    return st.session_state.pop("tmc_output", None) is not None


_WORKFLOW_SOURCE_UNSET = object()


def _workflow_state_for_mode(mode: str) -> WorkflowState | None:
    states = st.session_state.get(WORKFLOW_STATE_KEY)
    if not isinstance(states, dict):
        return None
    state = states.get(mode)
    return state if isinstance(state, WorkflowState) else None


def _store_workflow_state(state: WorkflowState) -> None:
    states = st.session_state.get(WORKFLOW_STATE_KEY)
    states = dict(states) if isinstance(states, dict) else {}
    states[state.mode] = state
    st.session_state[WORKFLOW_STATE_KEY] = states


def _workflow_export_payload(
    setup: dict[str, object] | None,
    *,
    export_mode: str | None,
    source_file_name: str | None = None,
    extra: dict[str, object] | None = None,
) -> dict[str, object]:
    setup = setup or {}
    payload: dict[str, object] = {
        "metadata": {field: setup.get(field, "") for field in WORKFLOW_EXPORT_METADATA_FIELDS},
        "export_mode": str(export_mode or ""),
        "source_file_name": str(source_file_name or ""),
        "template_version": TEMPLATE_VERSION,
        "template_name": Path(DEFAULT_TEMPLATE_PATH).name,
        "template_map_name": Path(DEFAULT_TEMPLATE_MAP_PATH).name,
        "template_path": str(DEFAULT_TEMPLATE_PATH),
        "template_map_path": str(DEFAULT_TEMPLATE_MAP_PATH),
        "use_template_report_layout": _use_template_layout_for_export(export_mode),
    }
    if extra:
        payload["extra"] = extra
    return payload


def _analysis_result_revision(revisions: WorkflowRevisions, marker: object) -> str | None:
    if marker is None:
        return None
    return analysis_result_fingerprint(
        {
            "source": revisions.source,
            "mapping": revisions.mapping,
            "analysis_config": revisions.analysis_config,
            "marker": marker,
        }
    )


def _single_workflow_revisions(
    *,
    source_bytes: bytes | None = None,
    source_revision: object = _WORKFLOW_SOURCE_UNSET,
    source_file_name: str | None,
    mapping: object,
    pce_factors: dict[str, float],
    peak_mode: str | None,
    peak_windows: dict[str, tuple[object, object]],
    movement_code_scheme: str,
    setup: dict[str, object],
    export_mode: str | None,
    confirmed_peaks: dict[str, str] | None,
    analysis_present: bool,
) -> WorkflowRevisions:
    source = (
        source_fingerprint(source_bytes)
        if source_revision is _WORKFLOW_SOURCE_UNSET
        else (str(source_revision) if source_revision else None)
    )
    revisions = WorkflowRevisions(
        source=source,
        mapping=mapping_fingerprint(mapping),
        analysis_config=analysis_config_fingerprint(
            pce_factors=pce_factors,
            peak_mode=peak_mode,
            peak_windows=peak_windows,
            movement_code_scheme=movement_code_scheme,
        ),
        review_decision=review_decision_fingerprint(confirmed_peaks),
        export_config=export_config_fingerprint(
            _workflow_export_payload(
                setup,
                export_mode=export_mode,
                source_file_name=source_file_name,
            )
        ),
    )
    if analysis_present:
        revisions = revisions.with_updates(analysis_result=_analysis_result_revision(revisions, "single"))
    return revisions


def _batch_workflow_revisions(
    *,
    uploads: list[object] | tuple[object, ...] | None,
    mapping_preset: dict[str, object] | None,
    pce_factors: dict[str, float],
    peak_mode: str | None,
    peak_windows: dict[str, tuple[object, object]],
    movement_code_scheme: str,
    metadata_rows: list[dict[str, object]],
    setup: dict[str, object],
    export_mode: str | None,
    confirmed_peaks: dict[str, dict[str, str]] | None,
    analysis_present: bool,
) -> WorkflowRevisions:
    content_revisions = tuple(source_fingerprint(file.getvalue()) for file in uploads or [])
    source = semantic_fingerprint(content_revisions) if content_revisions else None
    mapping = semantic_fingerprint(
        {
            "movement_code_scheme": movement_code_scheme,
            "mapping_rows": (mapping_preset or {}).get("mapping_rows", []),
        }
    )
    revisions = WorkflowRevisions(
        source=source,
        mapping=mapping,
        analysis_config=analysis_config_fingerprint(
            pce_factors=pce_factors,
            peak_mode=peak_mode,
            peak_windows=peak_windows,
            movement_code_scheme=movement_code_scheme,
        ),
        review_decision=review_decision_fingerprint(confirmed_peaks),
        export_config=export_config_fingerprint(
            _workflow_export_payload(
                setup,
                export_mode=export_mode,
                extra={
                    "metadata_rows": metadata_rows,
                    "preset_name": str((mapping_preset or {}).get("preset_name", "")),
                    "source_file_names": [str(getattr(file, "name", "")) for file in uploads or []],
                },
            )
        ),
    )
    if analysis_present:
        revisions = revisions.with_updates(analysis_result=_analysis_result_revision(revisions, "batch"))
    return revisions


def _clear_single_review_state() -> None:
    for key in (
        "am_peak_period_select",
        "pm_peak_period_select",
        "tmc_confirmed_am_peak_start",
        "tmc_confirmed_am_peak_end",
        "tmc_confirmed_pm_peak_start",
        "tmc_confirmed_pm_peak_end",
        "tmc_loaded_confirmed_peaks",
    ):
        st.session_state.pop(key, None)


def _apply_workflow_transition(mode: str, transition: WorkflowTransition) -> None:
    """Apply pure transition decisions to legacy session-state adapters."""

    if not transition.engineering_state_changed:
        return

    if mode == WORKFLOW_SINGLE_MODE:
        if transition.analysis_invalidated:
            if st.session_state.get("tmc_processed") is not None:
                st.session_state["tmc_pce_results_stale"] = True
            _clear_single_review_state()
        elif transition.review_invalidated:
            _clear_single_review_state()
        if transition.export_invalidated:
            _clear_single_export()
        return

    has_analysis = st.session_state.get("tmc_batch_analysis_result") is not None
    has_export = st.session_state.get("tmc_batch_export_result") is not None
    if transition.analysis_invalidated:
        if has_analysis:
            st.session_state["tmc_batch_stale"] = True
        st.session_state["tmc_batch_confirmed_peaks"] = {}
    elif transition.review_invalidated:
        st.session_state["tmc_batch_confirmed_peaks"] = {}
    if transition.analysis_invalidated or transition.review_invalidated or transition.export_invalidated:
        if has_analysis or has_export:
            st.session_state["tmc_batch_export_stale"] = True
        st.session_state.pop("tmc_batch_export_result", None)


def _sync_workflow_contract(
    mode: str,
    revisions: WorkflowRevisions,
    readiness: WorkflowReadiness,
) -> WorkflowTransition:
    previous_state = _workflow_state_for_mode(mode)
    transition = transition_workflow(
        previous_state.revisions if previous_state is not None else None,
        revisions,
    )
    _apply_workflow_transition(mode, transition)
    if transition.analysis_invalidated:
        revisions = revisions.with_updates(analysis_result=None)
    _store_workflow_state(
        WorkflowState(
            mode=mode,
            revisions=revisions,
            readiness=readiness_after_transition(readiness, transition),
        )
    )
    return transition


def _record_workflow_state(mode: str, revisions: WorkflowRevisions, readiness: WorkflowReadiness) -> None:
    _store_workflow_state(WorkflowState(mode=mode, revisions=revisions, readiness=readiness))


def apply_batch_export_mode_change(selected_mode: str, previous_mode: str | None) -> bool:
    set_active_tab("ส่งออก")
    if selected_mode == previous_mode:
        return False
    st.session_state["tmc_batch_export_mode"] = selected_mode
    _mark_batch_export_stale_now()
    return True


def _use_template_layout_for_export(export_mode: str | None) -> bool:
    return export_mode == EXCEL_TEMPLATE_EXPORT_MODE or str(export_mode or "").startswith(BATCH_EXCEL_TEMPLATE_EXPORT_MODE)


def _use_excel_native_charts_for_export(export_mode: str | None, excel_com_status: ExcelComStatus) -> bool:
    return bool(getattr(excel_com_status, "available", False) and _use_template_layout_for_export(export_mode))


def _movement_scheme(value: str | None) -> str:
    return normalize_movement_code_scheme(value or MOVEMENT_SCHEME_V1)


def _scheme_select_label(scheme: str) -> str:
    return movement_scheme_display_label(scheme)


def _scheme_status(scheme: str, status: str) -> str:
    return movement_scheme_status_label(scheme, status)


def _is_v2_scheme(value: str | None) -> bool:
    return _movement_scheme(value) == MOVEMENT_SCHEME_V2


def _is_v2_result(result: object | None) -> bool:
    return _is_v2_scheme(str(getattr(result, "movement_code_scheme", MOVEMENT_SCHEME_V1)))


def _single_file_processing_block_reason(movement_code_scheme: str) -> str:
    return "" if _is_v2_scheme(movement_code_scheme) else mapping_processing_block_reason(movement_code_scheme)


def _single_file_mapping_issues(
    detected_sheet_names: list[str],
    mapping: pd.DataFrame,
    movement_code_scheme: str,
) -> pd.DataFrame:
    return validate_mapping_for_processing_by_scheme(detected_sheet_names, mapping, movement_code_scheme)


def _movement_code_options_for_scheme(mapping: pd.DataFrame, movement_code_scheme: str) -> list[str]:
    base_options = APPROACH_MOVEMENT_CODES if _is_v2_scheme(movement_code_scheme) else MOVEMENT_CODE_OPTIONS
    values = mapping["movement_code"] if "movement_code" in mapping else None
    return selectbox_options_with_existing_values(["", *base_options], values)


MAPPING_SOURCE_DEFAULT_PREVIEW = "default_preview"
MAPPING_SOURCE_USER_EDITOR = "user_editor"
MAPPING_SOURCE_MAPPING_EXCEL = "mapping_excel"
MAPPING_SOURCE_MAPPING_PRESET = "mapping_preset"
MAPPING_SOURCE_PROJECT_SESSION = "project_session"
MAPPING_COMMITTED_SOURCES = {
    MAPPING_SOURCE_USER_EDITOR,
    MAPPING_SOURCE_MAPPING_EXCEL,
    MAPPING_SOURCE_MAPPING_PRESET,
    MAPPING_SOURCE_PROJECT_SESSION,
}
MAPPING_SCHEME_LOCK_CAPTION = "มีรายการ Mapping อยู่แล้ว หากต้องการเปลี่ยนระบบรหัส Movement กรุณาล้าง Mapping หรือโหลดไฟล์ Mapping ใหม่"


def _mapping_source(state: MutableMapping[str, object] | None = None) -> str:
    source_state = st.session_state if state is None else state
    source = str(source_state.get("tmc_mapping_source") or "").strip()
    return source or MAPPING_SOURCE_DEFAULT_PREVIEW


def _mapping_rows_are_committed(state: MutableMapping[str, object] | None = None) -> bool:
    source_state = st.session_state if state is None else state
    if not bool(source_state.get("mapping_table")):
        return False
    source = _mapping_source(source_state)
    if source in MAPPING_COMMITTED_SOURCES:
        return True
    if source == MAPPING_SOURCE_DEFAULT_PREVIEW:
        return False
    return bool(
        source_state.get("tmc_mapping_table_from_session")
        or source_state.get("tmc_mapping_upload_identity")
        or source_state.get("tmc_mapping_preset_upload_identity")
        or source_state.get("tmc_mapping_preset_apply_info")
    )


def _set_mapping_source(source: str, state: MutableMapping[str, object] | None = None) -> None:
    source_state = st.session_state if state is None else state
    source_state["tmc_mapping_source"] = source


def _clear_mapping_for_scheme_change(state: MutableMapping[str, object] | None = None) -> None:
    source_state = st.session_state if state is None else state
    if source_state.get("tmc_mapping_upload_identity"):
        source_state["tmc_mapping_upload_ignored_identity"] = source_state.get("tmc_mapping_upload_identity")
    if source_state.get("tmc_mapping_preset_upload_identity"):
        source_state["tmc_mapping_preset_ignored_identity"] = source_state.get("tmc_mapping_preset_upload_identity")
    for key in (
        "mapping_table",
        "tmc_loaded_mapping_code_scheme",
        "tmc_mapping_upload_identity",
        "tmc_mapping_preset_upload_identity",
        "tmc_mapping_preset_apply_info",
        "tmc_mapping_preset_warnings",
    ):
        source_state.pop(key, None)
    source_state["tmc_mapping_table_from_session"] = False
    source_state["tmc_mapping_source"] = MAPPING_SOURCE_DEFAULT_PREVIEW
    source_state["mapping_editor_version"] = int(source_state.get("mapping_editor_version", 0) or 0) + 1


def _mapping_editor_changed(before: pd.DataFrame, after: pd.DataFrame) -> bool:
    left = pd.DataFrame(before).reset_index(drop=True)
    right = pd.DataFrame(after).reset_index(drop=True)
    columns = list(dict.fromkeys([*left.columns.tolist(), *right.columns.tolist()]))
    left = left.reindex(columns=columns).fillna("").astype(str)
    right = right.reindex(columns=columns).fillna("").astype(str)
    return left.to_dict("records") != right.to_dict("records")


def _mapping_editor_labels(movement_code_scheme: str) -> dict[str, str]:
    common_labels = {
        "raw_sheet": "Sheet ต้นทาง",
        "raw_direction": "ทิศ/stream ต้นทาง",
        "source_stream": "source_stream",
        "movement_code": "รหัส movement",
        "raw_movement_label": "ป้ายแสดงผล",
        "include_in_report": "แสดงในรายงาน",
        "include_in_peak": "ใช้คำนวณ Peak",
    }
    if _is_v2_scheme(movement_code_scheme):
        return {
            **common_labels,
            "approach_direction": "ทิศทางการเดินทาง",
            "movement_type": "ประเภท movement",
        }
    return common_labels


def _movement_code_help_text(movement_code_scheme: str) -> str:
    if _is_v2_scheme(movement_code_scheme):
        return "รหัสแบบทิศทางการเดินทาง–รูปแบบการเคลื่อนที่ เช่น NT, WT, ET"
    return "รหัสแบบขาเข้า–ขาออก เช่น NS, WE, EW"


def _mapping_readiness_status(mapping: pd.DataFrame, movement_code_scheme: str, blocking_issue_count: int) -> str:
    if mapping.empty:
        return _scheme_status(movement_code_scheme, "รอกำหนดทิศทาง")
    if blocking_issue_count:
        return _scheme_status(movement_code_scheme, "ตรวจพบรหัสไม่ถูกต้อง")
    return _scheme_status(movement_code_scheme, "พร้อมใช้งาน")


def _process_single_file_for_ui(
    *,
    raw_sheets: dict[str, pd.DataFrame],
    mapping: pd.DataFrame,
    setup: dict[str, object],
    detected_sheets: list[str],
    peak_mode: str,
    peak_windows: dict[str, tuple[str, str]],
    pce_factors: dict[str, float],
) -> object:
    scheme = _movement_scheme(str(setup.get("movement_code_scheme") or MOVEMENT_SCHEME_V1))
    if scheme == MOVEMENT_SCHEME_V2:
        mapping = normalize_mapping_for_scheme(mapping, MOVEMENT_SCHEME_V2)
        return process_tmc_dry_run_v2(
            raw_sheets=raw_sheets,
            mapping=mapping,
            setup={**setup, "movement_code_scheme": MOVEMENT_SCHEME_V2},
            detected_sheets=detected_sheets,
            peak_mode=peak_mode,
            peak_windows=peak_windows,
            pce_factors=pce_factors,
        )
    mapping = normalize_mapping_for_scheme(mapping, MOVEMENT_SCHEME_V1)
    return process_tmc(
        raw_sheets=raw_sheets,
        mapping=mapping,
        setup={**setup, "movement_code_scheme": MOVEMENT_SCHEME_V1},
        detected_sheets=detected_sheets,
        peak_mode=peak_mode,
        peak_windows=peak_windows,
        pce_factors=pce_factors,
        generate_workbook=False,
    )


def _hourly_movement_for_ui(result: object | None, mapping: pd.DataFrame) -> pd.DataFrame:
    if result is None:
        return pd.DataFrame()
    if _is_v2_result(result):
        return getattr(result, "hourly_movement_pcu", pd.DataFrame()).copy()
    return hourly_movement_pcu(result.normalized, mapping)


def _v2_movement_code_reference_frame() -> pd.DataFrame:
    rows = []
    for code in APPROACH_MOVEMENT_CODES:
        rows.append(
            {
                "movement_code_scheme": MOVEMENT_SCHEME_V2,
                "movement_code": code,
                "approach_direction": code[0],
                "movement_type": code[1],
            }
        )
    return pd.DataFrame(rows)


def _v2_native_template_block_reason(
    export_mode: str | None,
    use_template_report_layout: bool,
    use_excel_com_native_charts: bool,
) -> str:
    if (export_mode == EXCEL_TEMPLATE_EXPORT_MODE or use_template_report_layout) and not use_excel_com_native_charts:
        return V2_EXCEL_TEMPLATE_MODE_BLOCK_MESSAGE
    return ""


def _export_single_file_for_ui(
    *,
    result: object,
    mapping: pd.DataFrame,
    setup: dict[str, object],
    export_mode: str,
    use_template_report_layout: bool,
    use_excel_com_native_charts: bool,
    source_file_name: str,
    generated_at: str,
) -> bytes:
    if _is_v2_result(result):
        block_reason = _v2_native_template_block_reason(export_mode, use_template_report_layout, use_excel_com_native_charts)
        if block_reason:
            raise ValueError(block_reason)
        if use_template_report_layout:
            return export_v2_template_workbook_com(
                result,
                setup={**setup, "movement_code_scheme": MOVEMENT_SCHEME_V2},
                mapping=mapping,
                export_mode=export_mode,
                source_file_name=source_file_name,
                generated_at=generated_at,
            )
        return export_v2_generated_workbook(
            result,
            setup={**setup, "movement_code_scheme": MOVEMENT_SCHEME_V2},
            mapping=mapping,
            export_mode=export_mode,
            source_file_name=source_file_name,
            generated_at=generated_at,
        )
    raise ValueError("v1 export continues to use the existing process_tmc export path.")


def _excel_com_status_fields(status: ExcelComStatus) -> dict[str, object]:
    return {
        "python_executable": sys.executable,
        "probed_at": getattr(status, "probed_at", "") or "",
        "available": bool(getattr(status, "available", False)),
        "reason": getattr(status, "reason", "UNKNOWN_STATUS"),
        "excel_version": getattr(status, "version", "") or "",
        "detail": getattr(status, "detail", "") or "",
        "pywintypes_file": getattr(status, "pywintypes_file", "") or "",
        "pywintypes_error": getattr(status, "pywintypes_error", "") or "",
        "pywin32_system32": getattr(status, "pywin32_system32", "") or "",
        "sys_path_win32_entries": list(getattr(status, "sys_path_matches", ()) or ()),
    }


def _render_excel_com_status(status: ExcelComStatus) -> None:
    try:
        fields = _excel_com_status_fields(status)
        st.write(fields)
        if getattr(status, "available", False):
            st.success(f"Excel COM พร้อมใช้งาน — Excel version: {getattr(status, 'version', '')}")
        else:
            st.warning(f"Excel COM ไม่พร้อมใช้งาน: {fields['reason']}")
            if fields["detail"]:
                st.caption(fields["detail"])
    except Exception as exc:  # pragma: no cover - UI guardrail
        st.warning(f"ไม่สามารถแสดงรายละเอียด Excel COM ได้: {exc}")


def _inject_global_css() -> None:
    st.markdown(
        """
        <style>
        :root {
            --tmc-bg: #F8F6F1;
            --tmc-surface: #FFFEFA;
            --tmc-surface-muted: #F1EFE8;
            --tmc-surface-raised: #FFFFFF;
            --tmc-sidebar: #EEEAE0;
            --tmc-sidebar-panel: #F8F6F1;

            --tmc-text: #151713;
            --tmc-text-muted: #5F625B;
            --tmc-text-soft: #73776E;

            --tmc-border: #DDD8CE;
            --tmc-border-soft: #ECE7DD;
            --tmc-divider: #E6E1D8;

            --tmc-primary: #0E4A2A;
            --tmc-primary-hover: #0A3B21;
            --tmc-primary-soft: #E8EFE7;
            --tmc-primary-subtle: #F1F6F0;

            --tmc-bronze: #B57A22;
            --tmc-bronze-soft: #F7EAD4;
            --tmc-bronze-subtle: #FCF6EA;

            --tmc-success: #0E6B3A;
            --tmc-success-soft: #E7F2EA;
            --tmc-warning: #B56A00;
            --tmc-warning-soft: #FAEBD2;
            --tmc-danger: #A23B2A;
            --tmc-danger-soft: #F8E7E3;
            --tmc-info: #2F5F8F;
            --tmc-info-soft: #E7EEF6;

            --tmc-font-display: Georgia, "Times New Roman", serif;
            --tmc-font-ui: "Noto Sans Thai", "IBM Plex Sans Thai", "Segoe UI", Tahoma, sans-serif;
            --tmc-font-mono: "IBM Plex Mono", "Roboto Mono", Consolas, monospace;
            --tmc-font-xs: 0.75rem;
            --tmc-font-sm: 0.875rem;
            --tmc-font-md: 0.96rem;
            --tmc-font-lg: 1.12rem;
            --tmc-font-xl: 2.15rem;

            --tmc-space-1: 0.25rem;
            --tmc-space-2: 0.5rem;
            --tmc-space-3: 0.75rem;
            --tmc-space-4: 1rem;
            --tmc-space-5: 1.25rem;
            --tmc-space-6: 1.5rem;

            --tmc-radius-sm: 6px;
            --tmc-radius-md: 8px;
            --tmc-radius-lg: 8px;
            --tmc-radius-xl: 10px;
            --tmc-shadow-subtle: 0 1px 2px rgba(21, 23, 19, 0.04);
            --tmc-shadow-soft: 0 8px 24px rgba(21, 23, 19, 0.06);
        }
        html, body, [class*="css"] {
            font-family: var(--tmc-font-ui);
        }
        .stApp {
            background: var(--tmc-bg);
            color: var(--tmc-text);
            font-family: var(--tmc-font-ui);
        }
        .block-container {
            max-width: 1240px;
            padding-top: 1.15rem;
            padding-bottom: 2.5rem;
            padding-left: 2rem;
            padding-right: 2rem;
        }
        .tmc-header {
            padding: var(--tmc-space-1) 0 var(--tmc-space-3) 0;
            border-bottom: 1px solid var(--tmc-divider);
            margin-bottom: var(--tmc-space-3);
        }
        .tmc-title {
            color: var(--tmc-text);
            font-size: 2rem;
            font-family: var(--tmc-font-ui);
            font-weight: 650;
            letter-spacing: 0;
            margin: 0;
            line-height: 1.15;
        }
        .tmc-subtitle {
            color: var(--tmc-text-muted);
            font-size: var(--tmc-font-md);
            margin-top: 0.35rem;
            line-height: 1.5;
            max-width: 760px;
        }
        .tmc-topbar {
            display: grid;
            grid-template-columns: repeat(5, minmax(0, 1fr));
            gap: 0;
            background: var(--tmc-surface);
            border: 1px solid var(--tmc-border);
            border-radius: var(--tmc-radius-lg);
            margin: 0 0 var(--tmc-space-3) 0;
            overflow: hidden;
            box-shadow: var(--tmc-shadow-subtle);
        }
        .tmc-topbar-item {
            min-width: 0;
            padding: 0.62rem 0.78rem;
            border-right: 1px solid var(--tmc-divider);
        }
        .tmc-topbar-item:last-child {
            border-right: 0;
        }
        .tmc-topbar-label {
            color: var(--tmc-text-soft);
            font-size: 0.68rem;
            font-weight: 700;
            line-height: 1.2;
            text-transform: uppercase;
        }
        .tmc-topbar-value {
            color: var(--tmc-text);
            font-size: var(--tmc-font-sm);
            font-weight: 650;
            line-height: 1.28;
            margin-top: 0.18rem;
            overflow-wrap: anywhere;
        }
        .tmc-topbar-note {
            color: var(--tmc-text-soft);
            font-size: 0.7rem;
            line-height: 1.25;
            margin-top: 0.08rem;
            overflow-wrap: anywhere;
        }
        .tmc-workflow {
            color: var(--tmc-text-soft);
            font-size: var(--tmc-font-sm);
            margin-top: var(--tmc-space-2);
        }
        .tmc-workflow-shell {
            background: var(--tmc-surface);
            border: 1px solid var(--tmc-border);
            border-radius: var(--tmc-radius-lg);
            padding: 0.72rem var(--tmc-space-4) 0.78rem;
            margin: 0 0 var(--tmc-space-3) 0;
            box-shadow: var(--tmc-shadow-subtle);
        }
        .tmc-workflow-shell-head {
            display: flex;
            align-items: baseline;
            justify-content: space-between;
            gap: var(--tmc-space-3);
            margin-bottom: var(--tmc-space-2);
        }
        .tmc-workflow-mode {
            color: var(--tmc-text);
            font-size: var(--tmc-font-sm);
            font-weight: 650;
            line-height: 1.25;
            text-transform: none;
        }
        .tmc-workflow-next {
            color: var(--tmc-bronze);
            font-size: var(--tmc-font-xs);
            font-weight: 650;
            line-height: 1.35;
            text-align: right;
            overflow-wrap: anywhere;
        }
        .tmc-status-grid {
            display: grid;
            grid-template-columns: repeat(3, minmax(0, 1fr));
            gap: var(--tmc-space-3);
            margin: var(--tmc-space-2) 0 var(--tmc-space-5) 0;
        }
        .tmc-card {
            background: var(--tmc-surface-raised);
            border: 1px solid var(--tmc-border);
            border-radius: var(--tmc-radius-lg);
            padding: var(--tmc-space-3) var(--tmc-space-4);
            box-sizing: border-box;
            box-shadow: var(--tmc-shadow-subtle);
        }
        .tmc-status-card {
            min-height: 78px;
            display: flex;
            flex-direction: column;
            justify-content: center;
            gap: 0.12rem;
            overflow: hidden;
            min-width: 0;
        }
        .tmc-card-label {
            color: var(--tmc-text-muted);
            font-size: var(--tmc-font-xs);
            font-weight: 600;
            letter-spacing: 0;
            line-height: 1.25;
        }
        .tmc-card-value {
            color: var(--tmc-text);
            font-size: var(--tmc-font-md);
            font-weight: 650;
            line-height: 1.28;
            overflow-wrap: anywhere;
        }
        .tmc-card-note {
            color: var(--tmc-text-soft);
            font-size: var(--tmc-font-xs);
            line-height: 1.25;
            overflow-wrap: anywhere;
        }
        .tmc-panel {
            background: var(--tmc-surface);
            border: 1px solid var(--tmc-border);
            border-radius: var(--tmc-radius-lg);
            padding: var(--tmc-space-4);
            margin: var(--tmc-space-2) 0 var(--tmc-space-4) 0;
            box-shadow: var(--tmc-shadow-subtle);
        }
        .tmc-panel-title {
            color: var(--tmc-text);
            font-size: var(--tmc-font-md);
            font-weight: 650;
            line-height: 1.3;
            margin-bottom: 0.12rem;
        }
        .tmc-panel-description {
            color: var(--tmc-text-muted);
            font-size: var(--tmc-font-sm);
            line-height: 1.4;
        }
        .tmc-kpi-grid {
            display: grid;
            grid-template-columns: repeat(4, minmax(0, 1fr));
            gap: var(--tmc-space-3);
            margin: var(--tmc-space-2) 0 var(--tmc-space-4) 0;
        }
        .tmc-kpi-card {
            min-height: 86px;
            display: flex;
            flex-direction: column;
            justify-content: center;
            gap: 0.1rem;
            overflow: hidden;
        }
        .tmc-kpi-value {
            color: var(--tmc-text);
            font-size: 1.32rem;
            font-weight: 700;
            line-height: 1.18;
            font-variant-numeric: tabular-nums;
            overflow-wrap: anywhere;
        }
        .tmc-kpi-unit {
            color: var(--tmc-text-muted);
            font-size: var(--tmc-font-xs);
            font-weight: 600;
            margin-left: 0.2rem;
        }
        .tmc-kpi-detail {
            color: var(--tmc-text-soft);
            font-size: var(--tmc-font-xs);
            line-height: 1.3;
            overflow-wrap: anywhere;
        }
        .tmc-chip {
            display: inline-flex;
            align-items: center;
            max-width: 100%;
            border: 1px solid var(--tmc-border-soft);
            border-radius: var(--tmc-radius-sm);
            padding: 0.1rem 0.46rem;
            font-size: var(--tmc-font-xs);
            font-weight: 650;
            line-height: 1.35;
            white-space: normal;
            overflow-wrap: anywhere;
        }
        .tmc-chip-success {
            color: #0A4F2A;
            background: var(--tmc-success-soft);
            border-color: #cfe3d4;
        }
        .tmc-chip-warning {
            color: #8A5100;
            background: var(--tmc-warning-soft);
            border-color: var(--tmc-bronze-soft);
        }
        .tmc-chip-danger {
            color: var(--tmc-danger);
            background: var(--tmc-danger-soft);
            border-color: #efd0c8;
        }
        .tmc-chip-info {
            color: var(--tmc-info);
            background: var(--tmc-info-soft);
            border-color: #d0ddea;
        }
        .tmc-chip-neutral {
            color: var(--tmc-text-muted);
            background: var(--tmc-surface-muted);
            border-color: var(--tmc-border);
        }
        .tmc-alert {
            border: 1px solid var(--tmc-border-soft);
            border-left-width: 3px;
            border-radius: var(--tmc-radius-lg);
            padding: var(--tmc-space-2) var(--tmc-space-3);
            margin: var(--tmc-space-2) 0 var(--tmc-space-3) 0;
            color: var(--tmc-text);
            font-size: var(--tmc-font-sm);
            line-height: 1.45;
            background: var(--tmc-surface);
        }
        .tmc-alert-info {
            border-left-color: var(--tmc-info);
            background: var(--tmc-surface);
        }
        .tmc-alert-warning {
            border-left-color: var(--tmc-warning);
            background: var(--tmc-bronze-subtle);
        }
        .tmc-alert-error {
            border-left-color: var(--tmc-danger);
            background: var(--tmc-danger-soft);
        }
        .tmc-alert-success {
            border-left-color: var(--tmc-success);
            background: var(--tmc-success-soft);
        }
        .tmc-action-hint {
            color: var(--tmc-text-muted);
            font-size: var(--tmc-font-sm);
            line-height: 1.4;
            margin: -0.2rem 0 var(--tmc-space-2) 0;
        }
        .tmc-workflow-stepper {
            display: grid;
            grid-template-columns: repeat(6, minmax(0, 1fr));
            gap: 0;
            align-items: start;
            margin: 0.45rem 0 0.35rem;
        }
        .tmc-workflow-step {
            position: relative;
            display: grid;
            grid-template-rows: 1.22rem auto;
            justify-items: center;
            row-gap: 0.24rem;
            padding: 0.08rem var(--tmc-space-1) 0.12rem;
            min-width: 0;
        }
        .tmc-workflow-step-rail {
            position: relative;
            display: flex;
            align-items: center;
            justify-content: center;
            width: 100%;
            min-width: 0;
            z-index: 0;
        }
        .tmc-workflow-step-rail::before {
            content: "";
            position: absolute;
            left: calc(50% + 0.58rem);
            right: calc(-50% + 0.58rem);
            border-top: 1px solid var(--tmc-border);
            z-index: 0;
        }
        .tmc-workflow-step:last-child .tmc-workflow-step-rail::before {
            display: none;
        }
        .tmc-workflow-dot {
            position: relative;
            z-index: 2;
            width: 1.16rem;
            height: 1.16rem;
            border-radius: 999px;
            border: 1px solid var(--tmc-border);
            background: var(--tmc-surface);
            color: var(--tmc-text-soft);
            display: inline-flex;
            align-items: center;
            justify-content: center;
            font-size: 0.65rem;
            font-weight: 700;
            line-height: 1;
        }
        .tmc-workflow-step-copy {
            position: relative;
            z-index: 1;
            min-width: 0;
            width: 100%;
        }
        .tmc-workflow-step-label {
            color: var(--tmc-text-muted);
            font-size: 0.76rem;
            font-weight: 650;
            line-height: 1.25;
            position: relative;
            z-index: 1;
            text-align: center;
            overflow-wrap: anywhere;
            width: 100%;
        }
        .tmc-workflow-step-state {
            color: var(--tmc-text-soft);
            font-size: 0.67rem;
            line-height: 1.25;
            margin-top: 0.1rem;
            position: relative;
            z-index: 1;
            text-align: center;
            width: 100%;
        }
        .tmc-workflow-completed .tmc-workflow-dot {
            color: #ffffff;
            background: var(--tmc-success);
            border-color: var(--tmc-success);
        }
        .tmc-workflow-completed .tmc-workflow-step-rail::before {
            border-top-color: #aacdb5;
        }
        .tmc-workflow-ready .tmc-workflow-dot {
            color: var(--tmc-primary);
            background: var(--tmc-surface);
            border-color: var(--tmc-success);
            box-shadow: 0 0 0 2px var(--tmc-success-soft);
        }
        .tmc-workflow-ready .tmc-workflow-step-rail::before {
            border-top-color: #aacdb5;
        }
        .tmc-workflow-ready .tmc-workflow-step-label {
            color: var(--tmc-primary);
        }
        .tmc-workflow-active .tmc-workflow-dot {
            color: var(--tmc-surface);
            background: var(--tmc-primary);
            border-color: var(--tmc-primary);
        }
        .tmc-workflow-active .tmc-workflow-step-label {
            color: var(--tmc-primary);
        }
        .tmc-workflow-warning .tmc-workflow-dot {
            color: #ffffff;
            background: var(--tmc-bronze);
            border-color: var(--tmc-bronze);
        }
        .tmc-workflow-warning .tmc-workflow-step-label {
            color: var(--tmc-warning);
        }
        .tmc-workflow-summary {
            display: flex;
            flex-wrap: wrap;
            gap: 0.35rem;
            margin-top: 0.45rem;
            padding-top: 0.48rem;
            border-top: 1px solid var(--tmc-divider);
        }
        .tmc-workflow-summary .tmc-chip {
            font-weight: 600;
        }
        div[class*="st-key-workflow_tab_"] {
            margin-bottom: var(--tmc-space-4);
            border-bottom: 1px solid var(--tmc-border);
        }
        div[class*="st-key-workflow_tab_"] button {
            min-height: 2.15rem;
            border: 0 !important;
            border-bottom: 2px solid transparent !important;
            border-radius: 0 !important;
            background: transparent !important;
            color: var(--tmc-text-muted) !important;
            padding: 0.42rem 0.5rem 0.36rem !important;
            box-shadow: none !important;
        }
        div[class*="st-key-workflow_tab_"] button:hover,
        div[class*="st-key-workflow_tab_"] button:focus-visible {
            border-bottom-color: var(--tmc-primary-soft) !important;
            background: var(--tmc-surface-muted) !important;
            color: var(--tmc-text) !important;
        }
        div[class*="st-key-workflow_tab_"] button[kind="primary"] {
            border-bottom-color: var(--tmc-primary) !important;
            background: var(--tmc-primary-subtle) !important;
            color: var(--tmc-primary) !important;
            font-weight: 650;
        }
        div[class*="st-key-workflow_tab_"] button p {
            color: inherit !important;
            font-size: var(--tmc-font-sm);
            line-height: 1.25;
            overflow-wrap: anywhere;
        }
        .tmc-peak-am {
            border-top-color: var(--tmc-success);
            background: var(--tmc-primary-subtle);
        }
        .tmc-peak-pm {
            border-top-color: var(--tmc-bronze);
            background: var(--tmc-bronze-subtle);
        }
        .tmc-status-card .tmc-card-label,
        .tmc-status-card .tmc-card-value,
        .tmc-status-card .tmc-card-note {
            max-width: 100%;
            overflow-wrap: anywhere;
        }
        .tmc-status-gray { border-left: 3px solid var(--tmc-border); }
        .tmc-status-blue { border-left: 3px solid var(--tmc-primary); }
        .tmc-status-green { border-left: 3px solid var(--tmc-success); }
        .tmc-status-amber { border-left: 3px solid var(--tmc-bronze); }
        .tmc-status-red { border-left: 3px solid var(--tmc-danger); }

        .tmc-section-header {
            display: flex;
            gap: var(--tmc-space-3);
            align-items: flex-start;
            margin: 0 0 var(--tmc-space-3) 0;
        }
        .tmc-section-accent {
            width: 3px;
            min-height: 2.1rem;
            border-radius: var(--tmc-radius-sm);
            background: var(--tmc-primary);
            margin-top: 0.08rem;
        }
        .tmc-section-title {
            color: var(--tmc-text);
            font-size: 1.05rem;
            font-weight: 650;
            line-height: 1.25;
            margin: 0;
        }
        .tmc-section-description {
            color: var(--tmc-text-muted);
            font-size: var(--tmc-font-sm);
            line-height: 1.35;
            margin-top: 0.18rem;
        }
        .tmc-section {
            background: var(--tmc-surface);
            border: 1px solid var(--tmc-border);
            border-radius: var(--tmc-radius-lg);
            padding: var(--tmc-space-4);
            margin: var(--tmc-space-3) 0 var(--tmc-space-4) 0;
        }
        .tmc-empty-state {
            background: var(--tmc-surface);
            border: 1px dashed var(--tmc-border);
            border-radius: var(--tmc-radius-lg);
            padding: var(--tmc-space-5);
            color: var(--tmc-text-muted);
            margin: var(--tmc-space-3) 0 var(--tmc-space-4) 0;
        }
        .tmc-empty-title {
            color: var(--tmc-text);
            font-size: var(--tmc-font-md);
            font-weight: 650;
            margin-bottom: var(--tmc-space-1);
        }
        .tmc-empty-description {
            color: var(--tmc-text-muted);
            font-size: var(--tmc-font-sm);
            line-height: 1.45;
        }
        .tmc-peak-card {
            min-height: 116px;
            border-top: 3px solid var(--tmc-text-soft);
        }
        .tmc-peak-suggested {
            border-top-color: var(--tmc-bronze);
            background: var(--tmc-bronze-subtle);
        }
        .tmc-peak-confirmed {
            border-top-color: var(--tmc-primary);
            background: var(--tmc-primary-subtle);
        }
        .tmc-peak-topline {
            display: flex;
            justify-content: space-between;
            gap: var(--tmc-space-2);
            align-items: flex-start;
            margin-bottom: var(--tmc-space-2);
        }
        .tmc-peak-badge {
            border-radius: var(--tmc-radius-sm);
            padding: 0.1rem 0.48rem;
            font-size: 0.68rem;
            font-weight: 650;
            line-height: 1.4;
            white-space: nowrap;
        }
        .tmc-peak-suggested .tmc-peak-badge {
            color: var(--tmc-warning);
            background: var(--tmc-warning-soft);
        }
        .tmc-peak-confirmed .tmc-peak-badge {
            color: var(--tmc-primary-hover);
            background: var(--tmc-primary-soft);
            border: 1px solid var(--tmc-border-soft);
        }
        .tmc-peak-time {
            color: var(--tmc-text);
            font-size: 1.32rem;
            font-weight: 650;
            line-height: 1.25;
            overflow-wrap: anywhere;
        }
        .tmc-peak-pcu {
            color: var(--tmc-text-muted);
            font-size: var(--tmc-font-sm);
            margin-top: var(--tmc-space-1);
        }
        div[data-testid="stMetric"] {
            background: var(--tmc-surface-raised);
            border: 1px solid var(--tmc-border);
            border-radius: var(--tmc-radius-lg);
            padding: var(--tmc-space-3) var(--tmc-space-4);
            box-shadow: var(--tmc-shadow-subtle);
        }
        section[data-testid="stSidebar"] {
            background: var(--tmc-sidebar);
            border-right: 1px solid var(--tmc-border);
        }
        section[data-testid="stSidebar"] > div {
            padding-top: var(--tmc-space-4);
            padding-left: 1rem;
            padding-right: 1rem;
        }
        .tmc-sidebar-brand {
            border: 1px solid var(--tmc-border);
            border-radius: var(--tmc-radius-lg);
            background: var(--tmc-sidebar-panel);
            padding: 0.85rem 0.9rem;
            margin: 0 0 var(--tmc-space-4) 0;
            box-shadow: var(--tmc-shadow-subtle);
        }
        .tmc-sidebar-mark {
            color: var(--tmc-primary);
            font-size: 0.72rem;
            font-weight: 800;
            line-height: 1;
            text-transform: uppercase;
        }
        .tmc-sidebar-title {
            color: var(--tmc-text);
            font-size: 1.05rem;
            font-weight: 720;
            line-height: 1.2;
            margin-top: 0.28rem;
        }
        .tmc-sidebar-subtitle {
            color: var(--tmc-text-muted);
            font-size: 0.76rem;
            line-height: 1.35;
            margin-top: 0.22rem;
        }
        .tmc-sidebar-section {
            color: var(--tmc-primary);
            font-size: 0.72rem;
            font-weight: 800;
            line-height: 1.25;
            text-transform: uppercase;
            border-top: 1px solid var(--tmc-divider);
            padding-top: 0.78rem;
            margin: 0.35rem 0 0.2rem 0;
        }
        .tmc-sidebar-section:first-child {
            border-top: 0;
            padding-top: 0;
        }
        section[data-testid="stSidebar"] h2,
        section[data-testid="stSidebar"] h3 {
            color: var(--tmc-primary);
            font-weight: 650;
            letter-spacing: 0;
        }
        section[data-testid="stSidebar"] h2 {
            font-size: 1rem;
        }
        section[data-testid="stSidebar"] h3 {
            font-size: 0.88rem;
        }
        section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p {
            color: var(--tmc-text-muted);
            line-height: 1.45;
        }
        section[data-testid="stSidebar"] div[data-testid="stVerticalBlock"] {
            gap: 0.8rem;
        }
        div[data-testid="stVerticalBlockBorderWrapper"] {
            border-color: var(--tmc-border);
            border-radius: var(--tmc-radius-lg);
            box-shadow: var(--tmc-shadow-subtle);
        }
        div[data-testid="stTabs"] [data-baseweb="tab-list"] {
            gap: var(--tmc-space-2);
            border-bottom: 1px solid var(--tmc-border);
            background: transparent;
            margin-top: 0.15rem;
        }
        div[data-testid="stTabs"] [data-baseweb="tab"] {
            color: var(--tmc-text-muted);
            border-radius: 0;
            padding: 0.5rem 0.25rem 0.58rem 0.25rem;
            margin-bottom: -1px;
            border-bottom: 3px solid transparent;
            font-weight: 560;
            background: transparent;
        }
        div[data-testid="stTabs"] [data-baseweb="tab"][aria-selected="true"] {
            color: var(--tmc-primary);
            background: transparent;
            border-bottom-color: var(--tmc-primary);
            font-weight: 650;
        }
        div[data-testid="stTabs"] [data-baseweb="tab"]:hover {
            color: var(--tmc-primary);
            background: transparent;
        }
        div[data-testid="stTabs"] [data-baseweb="tab-highlight"] {
            background-color: transparent !important;
        }
        h2, h3, h4 {
            color: var(--tmc-text);
            letter-spacing: 0;
            font-weight: 650;
        }
        h4 {
            font-size: 1rem !important;
            margin: 0.2rem 0 0.5rem 0 !important;
        }
        div[data-baseweb="input"],
        div[data-baseweb="select"],
        div[data-baseweb="base-input"] {
            border-radius: var(--tmc-radius-sm);
        }
        div[data-baseweb="input"] input,
        div[data-baseweb="base-input"] input {
            min-height: 2.35rem;
        }
        div[data-testid="stTextInput"] label,
        div[data-testid="stSelectbox"] label,
        div[data-testid="stTimeInput"] label,
        div[data-testid="stFileUploader"] label,
        div[data-testid="stRadio"] label {
            color: var(--tmc-text);
            font-weight: 600;
        }
        div.stButton > button,
        div.stDownloadButton > button {
            border-radius: var(--tmc-radius-sm);
            border-color: var(--tmc-border);
            background: var(--tmc-surface-raised);
            color: var(--tmc-text);
            min-height: 2.35rem;
            box-shadow: none;
            font-weight: 600;
        }
        div.stButton > button:hover,
        div.stDownloadButton > button:hover {
            border-color: var(--tmc-primary);
            color: var(--tmc-primary);
        }
        div.stButton > button[kind="primary"],
        button[data-testid="stBaseButton-primary"] {
            background: var(--tmc-primary);
            border-color: var(--tmc-primary);
            color: #ffffff;
        }
        div.stButton > button[kind="primary"]:hover,
        button[data-testid="stBaseButton-primary"]:hover {
            background: var(--tmc-primary-hover);
            border-color: var(--tmc-primary-hover);
            color: #ffffff;
        }
        div.stButton > button:disabled,
        div.stDownloadButton > button:disabled,
        button[data-testid="stBaseButton-primary"]:disabled {
            background: var(--tmc-border) !important;
            border-color: var(--tmc-border) !important;
            color: var(--tmc-text-soft) !important;
            box-shadow: none;
        }
        .tmc-checklist {
            background: var(--tmc-surface-raised);
            border: 1px solid var(--tmc-border);
            border-radius: var(--tmc-radius-lg);
            box-shadow: var(--tmc-shadow-subtle);
            margin: var(--tmc-space-2) 0 var(--tmc-space-4) 0;
            overflow: hidden;
        }
        .tmc-check-item {
            display: flex;
            align-items: center;
            gap: var(--tmc-space-2);
            background: transparent;
            border-bottom: 1px solid var(--tmc-divider);
            padding: 0.44rem var(--tmc-space-3);
            color: var(--tmc-text);
            font-size: var(--tmc-font-sm);
        }
        .tmc-check-item:last-child {
            border-bottom: 0;
        }
        .tmc-check-ready { border-left: 3px solid var(--tmc-success); }
        .tmc-check-warn { border-left: 3px solid var(--tmc-warning); }
        .tmc-check-icon {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            width: 1.15rem;
            height: 1.15rem;
            border-radius: var(--tmc-radius-sm);
            flex: 0 0 auto;
            font-weight: 700;
            font-size: 0.72rem;
        }
        .tmc-check-ready .tmc-check-icon {
            color: var(--tmc-success);
            background: var(--tmc-success-soft);
        }
        .tmc-check-warn .tmc-check-icon {
            color: var(--tmc-warning);
            background: var(--tmc-warning-soft);
        }
        .tmc-check-body {
            flex: 1 1 auto;
            min-width: 0;
        }
        .tmc-check-label {
            font-weight: 600;
            line-height: 1.3;
        }
        .tmc-check-status {
            color: var(--tmc-text-muted);
            font-size: var(--tmc-font-xs);
            line-height: 1.3;
        }
        .tmc-check-detail {
            color: var(--tmc-text-muted);
            font-size: var(--tmc-font-xs);
            line-height: 1.35;
        }
        .tmc-mode-note {
            color: var(--tmc-text-muted);
            font-size: var(--tmc-font-sm);
            border: 1px solid var(--tmc-border);
            border-radius: var(--tmc-radius-lg);
            padding: var(--tmc-space-2) var(--tmc-space-3);
            margin-bottom: var(--tmc-space-3);
        }
        .tmc-mode-note-success {
            color: var(--tmc-primary-hover);
            background: var(--tmc-primary-soft);
            border-color: var(--tmc-border-soft);
            border-left: 3px solid var(--tmc-primary);
        }
        .tmc-mode-note-warning {
            color: var(--tmc-warning);
            background: var(--tmc-warning-soft);
            border-color: var(--tmc-bronze-soft);
            border-left: 3px solid var(--tmc-warning);
        }
        .tmc-sidebar-badge {
            background: var(--tmc-surface);
            border: 1px solid var(--tmc-border);
            border-radius: var(--tmc-radius-lg);
            padding: var(--tmc-space-2) var(--tmc-space-3);
            margin: var(--tmc-space-2) 0 var(--tmc-space-3) 0;
            font-size: var(--tmc-font-sm);
            color: var(--tmc-text);
            box-shadow: var(--tmc-shadow-subtle);
        }
        .tmc-sidebar-badge strong {
            color: var(--tmc-primary);
        }
        .tmc-sidebar-badge-success { border-left: 3px solid var(--tmc-success); }
        .tmc-sidebar-badge-warning { border-left: 3px solid var(--tmc-warning); }
        .tmc-version-stamp {
            color: var(--tmc-text-soft);
            font-size: var(--tmc-font-xs);
            line-height: 1.35;
            margin-top: var(--tmc-space-3);
        }
        div[data-testid="stAlert"] {
            border-radius: var(--tmc-radius-lg);
            background: transparent;
            color: var(--tmc-text);
            box-shadow: none;
        }
        div[data-testid="stAlert"] .stAlertContainer {
            background: var(--tmc-surface) !important;
            border: 1px solid var(--tmc-border-soft) !important;
            border-left: 3px solid var(--tmc-primary) !important;
            border-radius: var(--tmc-radius-lg) !important;
            color: var(--tmc-text) !important;
            box-shadow: none !important;
        }
        div[data-testid="stAlert"] [data-testid="stMarkdownContainer"] p {
            color: inherit;
        }
        div[data-testid="stDataFrame"],
        div[data-testid="stDataEditor"] {
            border-radius: var(--tmc-radius-lg);
            overflow: hidden;
            border: 1px solid var(--tmc-border-soft);
        }
        code, pre {
            font-family: var(--tmc-font-mono);
        }
        @media (max-width: 900px) {
            .block-container {
                padding-left: 1rem;
                padding-right: 1rem;
            }
            .tmc-topbar {
                grid-template-columns: repeat(2, minmax(0, 1fr));
            }
            .tmc-topbar-item {
                border-bottom: 1px solid var(--tmc-divider);
            }
            .tmc-topbar-item:nth-child(2n) {
                border-right: 0;
            }
            .tmc-status-grid {
                grid-template-columns: repeat(2, minmax(0, 1fr));
            }
            .tmc-kpi-grid,
            .tmc-workflow-stepper {
                grid-template-columns: repeat(2, minmax(0, 1fr));
                row-gap: var(--tmc-space-2);
            }
            .tmc-workflow-shell-head {
                align-items: flex-start;
                flex-direction: column;
            }
            .tmc-workflow-next {
                text-align: left;
            }
            .tmc-workflow-step-rail::before {
                display: none;
            }
        }
        @media (max-width: 640px) {
            .tmc-topbar {
                grid-template-columns: 1fr;
            }
            .tmc-topbar-item {
                border-right: 0;
            }
            .tmc-status-grid {
                grid-template-columns: 1fr;
            }
            .tmc-kpi-grid,
            .tmc-workflow-stepper {
                grid-template-columns: 1fr;
            }
            .tmc-title {
                font-size: 1.55rem;
            }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def _render_section_header(title: str, description: str = "") -> None:
    description_html = (
        f'<div class="tmc-section-description">{escape(description)}</div>' if description else ""
    )
    st.markdown(
        '<div class="tmc-section-header">'
        '<div class="tmc-section-accent"></div>'
        '<div>'
        f'<div class="tmc-section-title">{escape(title)}</div>'
        f"{description_html}"
        "</div>"
        "</div>",
        unsafe_allow_html=True,
    )


def _render_empty_state(title: str, description: str = "") -> None:
    description_html = (
        f'<div class="tmc-empty-description">{escape(description)}</div>' if description else ""
    )
    st.markdown(
        '<div class="tmc-empty-state">'
        f'<div class="tmc-empty-title">{escape(title)}</div>'
        f"{description_html}"
        "</div>",
        unsafe_allow_html=True,
    )


def _render_panel(title: str, description: str = "") -> None:
    description_html = (
        f'<div class="tmc-panel-description">{escape(description)}</div>' if description else ""
    )
    st.markdown(
        '<div class="tmc-panel">'
        f'<div class="tmc-panel-title">{escape(title)}</div>'
        f"{description_html}"
        "</div>",
        unsafe_allow_html=True,
    )


def _chip_kind_from_text(value: str) -> str:
    text = str(value).casefold()
    success_terms = {
        "success",
        "complete",
        "ready",
        "loaded",
        "ผ่าน",
        "สำเร็จ",
        "เสร็จสิ้น",
        "พร้อม",
        "พร้อมใช้งาน",
        "โหลดแล้ว",
        "กำหนดแล้ว",
        "เสร็จแล้ว",
        "วิเคราะห์แล้ว",
        "สร้างแล้ว",
    }
    warning_terms = {
        "warning",
        "warn",
        "needs review",
        "pending",
        "active",
        "ต้องตรวจสอบ",
        "ต้องวิเคราะห์ใหม่",
        "รอตรวจสอบ",
    }
    danger_terms = {"error", "failed", "fail", "danger", "ผิดพลาด", "ล้มเหลว", "ไม่สำเร็จ"}
    info_terms = {"info", "technical", "ข้อมูล"}
    if any(term in text for term in danger_terms):
        return "danger"
    if any(term in text for term in warning_terms):
        return "warning"
    if any(term in text for term in success_terms):
        return "success"
    if any(term in text for term in info_terms):
        return "info"
    return "neutral"


def _status_chip_html(label: str, kind: str = "neutral") -> str:
    allowed = {"success", "warning", "danger", "info", "neutral"}
    chip_kind = kind if kind in allowed else "neutral"
    return f'<span class="tmc-chip tmc-chip-{chip_kind}">{escape(label)}</span>'


def _render_status_chip(label: str, kind: str = "neutral") -> None:
    st.markdown(_status_chip_html(label, kind), unsafe_allow_html=True)


def _kpi_card_html(label: str, value: object, unit: str = "", detail: str = "", status: str | None = None) -> str:
    unit_html = f'<span class="tmc-kpi-unit">{escape(unit)}</span>' if unit else ""
    detail_html = f'<div class="tmc-kpi-detail">{escape(detail)}</div>' if detail else '<div class="tmc-kpi-detail">&nbsp;</div>'
    status_html = _status_chip_html(status, _chip_kind_from_text(status)) if status else ""
    return (
        '<div class="tmc-card tmc-kpi-card">'
        f'<div class="tmc-card-label">{escape(label)}</div>'
        f'<div class="tmc-kpi-value">{escape(str(value))}{unit_html}</div>'
        f"{detail_html}"
        f"{status_html}"
        "</div>"
    )


def _render_kpi_card(label: str, value: object, unit: str = "", detail: str = "", status: str | None = None) -> None:
    st.markdown(_kpi_card_html(label, value, unit, detail, status), unsafe_allow_html=True)


def _render_metric_strip(items: list[tuple[str, object, str, str] | tuple[str, object, str, str, str]], columns: int = 4) -> None:
    column_count = max(1, min(columns, 6))
    grid_style = f' style="grid-template-columns: repeat({column_count}, minmax(0, 1fr));"'
    html = ['<div class="tmc-kpi-grid"' + grid_style + ">"]
    for item in items:
        label, value, unit, detail, *rest = item
        status = rest[0] if rest else None
        html.append(_kpi_card_html(str(label), value, str(unit or ""), str(detail or ""), status))
    html.append("</div>")
    st.markdown("".join(html), unsafe_allow_html=True)


def _render_alert(message: str, kind: str = "info") -> None:
    allowed = {"info", "warning", "error", "success"}
    alert_kind = kind if kind in allowed else "info"
    st.markdown(
        f'<div class="tmc-alert tmc-alert-{alert_kind}">{escape(message)}</div>',
        unsafe_allow_html=True,
    )


def _render_action_hint(message: str) -> None:
    if message:
        st.markdown(f'<div class="tmc-action-hint">{escape(message)}</div>', unsafe_allow_html=True)


def _status_tone(status: str) -> str:
    text = str(status).casefold()
    if _chip_kind_from_text(status) == "success":
        return "green"
    if _chip_kind_from_text(status) == "danger":
        return "red"
    if _chip_kind_from_text(status) == "warning":
        return "amber"
    if _chip_kind_from_text(status) == "info":
        return "blue"
    return "gray"


def _status_card_html(label: str, status: str, note: str = "", tone: str | None = None) -> str:
    tone = tone or _status_tone(status)
    note_html = f'<div class="tmc-card-note">{escape(note)}</div>' if note else '<div class="tmc-card-note">&nbsp;</div>'
    chip_kind = {
        "green": "success",
        "amber": "warning",
        "red": "danger",
        "blue": "info",
    }.get(tone, _chip_kind_from_text(status))
    return (
        f'<div class="tmc-card tmc-status-card tmc-status-{tone}">'
        f'<div class="tmc-card-label">{escape(label)}</div>'
        f'<div class="tmc-card-value">{_status_chip_html(status, chip_kind)}</div>'
        f"{note_html}"
        "</div>"
    )


def _render_status_card(label: str, status: str, note: str = "") -> None:
    st.markdown(
        _status_card_html(label, status, note),
        unsafe_allow_html=True,
    )


def _render_sidebar_badge(title: str, detail: str, *, ready: bool) -> None:
    badge_class = "tmc-sidebar-badge-success" if ready else "tmc-sidebar-badge-warning"
    st.markdown(
        f'<div class="tmc-sidebar-badge {badge_class}">'
        f"<strong>{escape(title)}</strong><br>"
        f"<span>{escape(detail)}</span>"
        "</div>",
        unsafe_allow_html=True,
    )


def _render_version_stamp() -> None:
    st.markdown(
        f'<div class="tmc-version-stamp">'
        f"TMC Processor v{escape(APP_VERSION)}<br>"
        f"Template: {escape(TEMPLATE_VERSION)}"
        "</div>",
        unsafe_allow_html=True,
    )


def _render_sidebar_brand() -> None:
    st.markdown(
        '<div class="tmc-sidebar-brand">'
        '<div class="tmc-sidebar-mark">TMC Processor</div>'
        '<div class="tmc-sidebar-title">Engineering Workstation</div>'
        '<div class="tmc-sidebar-subtitle">Turning Movement Count processing and report control</div>'
        "</div>",
        unsafe_allow_html=True,
    )


def _render_sidebar_section(label: str) -> None:
    st.markdown(f'<div class="tmc-sidebar-section">{escape(label)}</div>', unsafe_allow_html=True)


def _render_app_header() -> None:
    st.markdown(
        """
        <div class="tmc-header">
            <h1 class="tmc-title">TMC Processor</h1>
            <div class="tmc-subtitle">ประมวลผลข้อมูล Turning Movement Count และจัดทำรายงาน Excel</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def _topbar_item(label: str, value: str, note: str = "") -> str:
    note_html = f'<div class="tmc-topbar-note">{escape(note)}</div>' if note else ""
    return (
        '<div class="tmc-topbar-item">'
        f'<div class="tmc-topbar-label">{escape(label)}</div>'
        f'<div class="tmc-topbar-value">{escape(value)}</div>'
        f"{note_html}"
        "</div>"
    )


def _render_top_status_bar(
    *,
    is_single_file_mode: bool,
    uploaded_name: str | None,
    uploaded_count: int,
    batch_mapping_ready: bool,
    export_mode: str | None,
    excel_com_status: ExcelComStatus,
) -> None:
    mode_value = "ไฟล์เดียว" if is_single_file_mode else "Batch"
    current_scheme = _current_mapping_scheme()
    workflow_state = _workflow_state_for_mode(
        WORKFLOW_SINGLE_MODE if is_single_file_mode else WORKFLOW_BATCH_MODE
    )
    if is_single_file_mode:
        source_ready = workflow_state.readiness.source if workflow_state is not None else bool(uploaded_name)
        mapping_rows = len(st.session_state.get("mapping_table") or [])
        mapping_ready = workflow_state.readiness.mapping if workflow_state is not None else bool(mapping_rows)
        source_value = "โหลดแล้ว" if source_ready else "ยังไม่มีไฟล์สำรวจ"
        source_note = "พร้อมกำหนด Mapping" if source_ready else "อัปโหลดจากแถบด้านซ้าย"
        mapping_value = "พร้อมใช้งาน" if mapping_ready else "ยังไม่พร้อม"
        mapping_note = f"{current_scheme} · {mapping_rows:,} แถว" if mapping_rows else f"{current_scheme} · รอกำหนดทิศทาง"
    else:
        source_ready = workflow_state.readiness.source if workflow_state is not None else bool(uploaded_count)
        mapping_ready = workflow_state.readiness.mapping if workflow_state is not None else batch_mapping_ready
        source_value = f"{uploaded_count:,} ไฟล์" if source_ready else "ยังไม่มีไฟล์ Batch"
        source_note = "โหลดไฟล์ Batch แล้ว" if source_ready else "อัปโหลดจากแถบด้านซ้าย"
        mapping_value = "พร้อมใช้งาน" if mapping_ready else "ยังไม่พร้อม"
        mapping_note = "Mapping Preset ใช้ร่วมกันทุกไฟล์" if batch_mapping_ready else "เปิด Mapping Preset"

    excel_value = "Excel COM พร้อม" if getattr(excel_com_status, "available", False) else "โหมดสำรอง PNG"
    excel_note = (
        f"Excel {excel_com_status.version}"
        if getattr(excel_com_status, "available", False) and getattr(excel_com_status, "version", "")
        else (str(getattr(excel_com_status, "reason", "")) or "COM unavailable")
    )
    export_note = export_mode or "รอเลือกโหมดส่งออก"

    html = (
        '<div class="tmc-topbar">'
        + _topbar_item("โหมดงาน", mode_value, export_note)
        + _topbar_item("ไฟล์สำรวจ", source_value, source_note)
        + _topbar_item("Mapping", mapping_value, mapping_note)
        + _topbar_item("เครื่องมือส่งออก", excel_value, excel_note)
        + _topbar_item("เวอร์ชัน", f"App v{APP_VERSION}", f"Template {TEMPLATE_VERSION}")
        + "</div>"
    )
    st.markdown(html, unsafe_allow_html=True)


def _workflow_step_statuses(*, is_single_file_mode: bool, uploaded: bool, batch_mapping_ready: bool = False) -> list[str]:
    mapping_ready = bool(st.session_state.get("mapping_table"))
    processed = "tmc_processed" in st.session_state
    output_ready = "tmc_output" in st.session_state
    batch_analysis_ready = st.session_state.get("tmc_batch_analysis_result") is not None
    batch_export_ready = st.session_state.get("tmc_batch_export_result") is not None
    if is_single_file_mode:
        done = [uploaded, uploaded, mapping_ready, processed, processed, output_ready]
    else:
        batch_files_ready = bool(st.session_state.get("tmc_batch_file_metadata_table"))
        preset_ready = batch_mapping_ready
        done = [uploaded, batch_files_ready, preset_ready, batch_analysis_ready, batch_analysis_ready, batch_export_ready]
    active_index = next((index for index, value in enumerate(done) if not value), len(done) - 1)
    return ["done" if value else ("active" if index == active_index else "pending") for index, value in enumerate(done)]


def _render_workflow_stepper(*, is_single_file_mode: bool, uploaded: bool, batch_mapping_ready: bool = False) -> None:
    labels = ["อัปโหลดไฟล์", "ตั้งค่างาน", "Mapping", "ประมวลผล", "Peak Review", "Export"]
    state_text = {"done": "พร้อม", "active": "ขั้นตอนปัจจุบัน", "pending": "รอดำเนินการ"}
    statuses = _workflow_step_statuses(
        is_single_file_mode=is_single_file_mode,
        uploaded=uploaded,
        batch_mapping_ready=batch_mapping_ready,
    )
    html = ['<div class="tmc-workflow-stepper">']
    for label, status in zip(labels, statuses):
        html.append(
            f'<div class="tmc-workflow-step tmc-workflow-{status}">'
            f'<div class="tmc-workflow-step-label">{escape(label)}</div>'
            f'<div class="tmc-workflow-step-state">{escape(state_text[status])}</div>'
            "</div>"
        )
    html.append("</div>")
    st.markdown("".join(html), unsafe_allow_html=True)


def _render_status_cards(
    *,
    uploaded_name: str | None,
    excel_com_status: ExcelComStatus,
    export_mode: str | None,
) -> None:
    mapping_rows = len(st.session_state.get("mapping_table") or [])
    confirmed = _confirmed_peaks_from_state()
    processed = "tmc_processed" in st.session_state
    output_ready = "tmc_output" in st.session_state
    excel_ready = bool(getattr(excel_com_status, "available", False)) or export_mode != EXCEL_TEMPLATE_EXPORT_MODE

    raw_status = "โหลดแล้ว" if uploaded_name else "ยังไม่ได้โหลด"
    session_status = "โหลดแล้ว" if st.session_state.get("tmc_loaded_project_session") else "ยังไม่ได้โหลด"
    mapping_status = "พร้อมใช้งาน" if mapping_rows else ("ต้องตรวจสอบ" if uploaded_name else "ยังไม่ได้โหลด")
    processing_status = "เสร็จสิ้น" if processed else ("ต้องตรวจสอบ" if mapping_rows else "ยังไม่ได้โหลด")
    peak_status = "กำหนดแล้ว" if all(
        confirmed.get(key) for key in ("am_peak_start", "am_peak_end", "pm_peak_start", "pm_peak_end")
    ) else ("ต้องตรวจสอบ" if processed else "ยังไม่ได้โหลด")
    export_status = "พร้อมใช้งาน" if output_ready or (processed and peak_status == "กำหนดแล้ว" and excel_ready) else "ยังไม่ได้โหลด"

    card_items = [
        ("ไฟล์สำรวจ", raw_status, uploaded_name or ""),
        ("Project Session", session_status, st.session_state.get("tmc_loaded_source_file_name", "")),
        ("การกำหนดทิศทาง", mapping_status, f"{mapping_rows:,} แถว" if mapping_rows else ""),
        ("การประมวลผล", processing_status, "พร้อมตรวจสอบกราฟ" if processed else ""),
        ("ช่วงเร่งด่วน", peak_status, ""),
        ("ความพร้อมส่งออก", export_status, "Excel COM" if export_mode == EXCEL_TEMPLATE_EXPORT_MODE else "PNG fallback"),
    ]
    st.markdown(
        '<div class="tmc-status-grid">'
        + "".join(_status_card_html(label, status, note) for label, status, note in card_items)
        + "</div>",
        unsafe_allow_html=True,
    )

def _workflow_status_label(status: str) -> str:
    return {
        "completed": "เสร็จแล้ว",
        "ready": "พร้อม",
        "active": "ขั้นตอนถัดไป",
        "pending": "รอดำเนินการ",
        "warning": "ต้องตรวจสอบ",
        "blocked": "ติดเงื่อนไข",
    }.get(status, "รอดำเนินการ")


def _workflow_summary_chip(label: str, value: str, kind: str = "neutral") -> str:
    return _status_chip_html(f"{label}: {value}", kind)


def _processed_result_from_state() -> object | None:
    processed = st.session_state.get("tmc_processed")
    if not processed or bool(st.session_state.get("tmc_pce_results_stale")):
        return None
    if isinstance(processed, dict):
        return processed.get("result")
    return processed


def _suggested_peaks_from_result(result: object | None) -> dict[str, str]:
    peaks = getattr(result, "peaks", None)
    if peaks is None:
        return {}
    am_start, am_end, _ = _peak_period_text(peaks, "AM")
    pm_start, pm_end, _ = _peak_period_text(peaks, "PM")
    return {
        "am_peak_start": am_start,
        "am_peak_end": am_end,
        "pm_peak_start": pm_start,
        "pm_peak_end": pm_end,
    }


def _single_effective_peak_state(result: object | None = None) -> dict[str, object]:
    result = _processed_result_from_state() if result is None else result
    if result is None:
        return {
            "ready": False,
            "source": "",
            "summary_text": "ยังไม่มีผลประมวลผล",
            "summary_kind": "neutral",
            "values": {},
        }

    suggested = _suggested_peaks_from_result(result)
    confirmed = _confirmed_peaks_from_state()
    values = dict(suggested)
    for key in ("am_peak_start", "am_peak_end", "pm_peak_start", "pm_peak_end"):
        if confirmed.get(key):
            values[key] = str(confirmed[key])
    ready = all(values.get(key) for key in ("am_peak_start", "am_peak_end", "pm_peak_start", "pm_peak_end"))
    if not ready:
        return {
            "ready": False,
            "source": "",
            "summary_text": "ขาด AM/PM Peak",
            "summary_kind": "warning",
            "values": values,
        }

    suggested_am = f"{suggested.get('am_peak_start', '')}-{suggested.get('am_peak_end', '')}".strip("-")
    suggested_pm = f"{suggested.get('pm_peak_start', '')}-{suggested.get('pm_peak_end', '')}".strip("-")
    effective_am = f"{values.get('am_peak_start', '')}-{values.get('am_peak_end', '')}".strip("-")
    effective_pm = f"{values.get('pm_peak_start', '')}-{values.get('pm_peak_end', '')}".strip("-")
    user_adjusted = bool(confirmed) and (effective_am != suggested_am or effective_pm != suggested_pm)
    loaded_peaks = bool(st.session_state.get("tmc_loaded_confirmed_peaks"))
    if user_adjusted:
        source = PEAK_SELECTION_USER_CONFIRMED
        summary_text = "กำหนดแล้ว"
    elif loaded_peaks:
        source = str(confirmed.get("peak_selection_source") or PEAK_SELECTION_USER_CONFIRMED)
        summary_text = "กำหนดแล้ว"
    else:
        source = PEAK_SELECTION_AUTO
        summary_text = "ใช้ค่าแนะนำ"
    return {
        "ready": True,
        "source": source,
        "summary_text": summary_text,
        "summary_kind": "success",
        "values": values,
    }


def _peak_range_label(values: dict[str, object], start_key: str, end_key: str) -> str:
    start = str(values.get(start_key) or "")
    end = str(values.get(end_key) or "")
    return f"{start}-{end}" if start and end else "-"


def _peak_source_label(source: object) -> str:
    source_text = str(source or "")
    if source_text == PEAK_SELECTION_USER_CONFIRMED:
        return "เลือกโดยผู้ใช้"
    if source_text == PEAK_SELECTION_AUTO:
        return "ระบบแนะนำ"
    return "ค่าเริ่มต้นของเทมเพลต"


def _render_effective_peak_export_confirmation(result: object | None, peak_state: dict[str, object]) -> None:
    suggested = _suggested_peaks_from_result(result)
    effective = dict(peak_state.get("values") or {})
    source_label = _peak_source_label(peak_state.get("source"))
    st.info(
        "\n".join(
            [
                f"Peak ที่ระบบแนะนำ: AM {_peak_range_label(suggested, 'am_peak_start', 'am_peak_end')} / PM {_peak_range_label(suggested, 'pm_peak_start', 'pm_peak_end')}",
                f"Peak ที่จะใช้ส่งออก: AM {_peak_range_label(effective, 'am_peak_start', 'am_peak_end')} / PM {_peak_range_label(effective, 'pm_peak_start', 'pm_peak_end')} ({source_label})",
            ]
        )
    )


def _readiness_payload(readiness: WorkflowReadiness) -> dict[str, bool]:
    return {
        "source": readiness.source,
        "mapping": readiness.mapping,
        "analysis": readiness.analysis,
        "review": readiness.review,
        "export": readiness.export,
    }


def _single_shell_from_authoritative_readiness(
    readiness: WorkflowReadiness,
    legacy_state: dict[str, object],
) -> dict[str, object]:
    """Render engineering stages from the stored UX-0 state, not legacy adapters."""

    steps = ["pending"] * 6
    steps[0] = "completed" if readiness.source else "active"
    steps[1] = list(legacy_state["steps"])[1]
    steps[2] = "completed" if readiness.mapping else ("active" if readiness.source else "pending")
    steps[3] = "completed" if readiness.analysis else ("active" if readiness.mapping else "pending")
    steps[4] = "ready" if readiness.review else ("active" if readiness.analysis else "pending")
    steps[5] = "completed" if readiness.export else ("active" if readiness.review else "pending")

    if not readiness.source:
        next_action = "เริ่มจากอัปโหลดไฟล์ TMC Excel ที่แถบด้านซ้าย"
    elif not readiness.mapping:
        next_action = "ตรวจสอบ Mapping ก่อนประมวลผล"
    elif not readiness.analysis:
        next_action = "ประมวลผลไฟล์หลัง Mapping พร้อมใช้งาน"
    elif not readiness.review:
        next_action = "กำหนดช่วง Peak ก่อนส่งออก"
    elif not readiness.export:
        next_action = "พร้อมส่งออกรายงาน"
    else:
        next_action = "สร้างรายงานแล้ว พร้อมดาวน์โหลดไฟล์"

    legacy_summary = list(legacy_state["summary"])
    summary = [
        (legacy_summary[0][0], "โหลดแล้ว" if readiness.source else "ยังไม่ได้โหลด", "success" if readiness.source else "neutral"),
        legacy_summary[1],
        (legacy_summary[2][0], "พร้อมใช้งาน" if readiness.mapping else "ยังไม่พร้อม", "success" if readiness.mapping else "neutral"),
        (legacy_summary[3][0], "เสร็จแล้ว" if readiness.analysis else "ยังไม่ได้ประมวลผล", "success" if readiness.analysis else "neutral"),
        (
            legacy_summary[4][0],
            "กำหนดแล้ว" if readiness.review else ("รอตรวจ Peak" if readiness.analysis else "ยังไม่มีผลประมวลผล"),
            "success" if readiness.review else ("warning" if readiness.analysis else "neutral"),
        ),
        (
            legacy_summary[5][0],
            "สร้างแล้ว" if readiness.export else ("รอสร้างรายงาน" if readiness.review else "ยังไม่พร้อม"),
            "success" if readiness.export else "neutral",
        ),
    ]
    return {"steps": steps, "summary": summary, "next_action": next_action, "readiness": _readiness_payload(readiness)}


def derive_single_workflow_state(uploaded_name: str | None, export_mode: str | None, excel_com_status: ExcelComStatus) -> dict[str, object]:
    mapping_rows = len(st.session_state.get("mapping_table") or [])
    mapping_frame = pd.DataFrame(st.session_state.get("mapping_table") or [])
    mapping_scheme = _current_mapping_scheme()
    scheme_issue_count = len(validate_mapping_scheme(mapping_frame, mapping_scheme)) if mapping_rows else 0
    processed = _processed_result_from_state() is not None
    output_ready = st.session_state.get("tmc_output") is not None
    peak_state = _single_effective_peak_state()
    peaks_ready = bool(peak_state["ready"])
    excel_ready = bool(getattr(excel_com_status, "available", False)) or export_mode != EXCEL_TEMPLATE_EXPORT_MODE
    preset_info = st.session_state.get("tmc_mapping_preset_apply_info") or {}
    mapping_needs_review = bool(uploaded_name and (not mapping_rows or int(preset_info.get("missing", 0) or 0) > 0))

    steps = ["pending"] * 6
    if uploaded_name:
        steps[0] = "completed"
        steps[1] = "completed"
    if mapping_rows and not mapping_needs_review and not scheme_issue_count:
        steps[2] = "completed"
    elif mapping_needs_review or scheme_issue_count:
        steps[2] = "warning"
    if processed:
        steps[3] = "completed"
    elif st.session_state.get("tmc_pce_results_stale"):
        steps[3] = "warning"
    if peaks_ready:
        steps[4] = "ready"
    elif processed:
        steps[4] = "warning"
    if output_ready:
        steps[5] = "completed"
    elif processed and peaks_ready and excel_ready:
        steps[5] = "ready"

    if not uploaded_name:
        steps[0] = "active"
        next_action = "เริ่มจากอัปโหลดไฟล์ TMC Excel ที่แถบด้านซ้าย"
    elif mapping_needs_review:
        next_action = "ตรวจสอบ Mapping ก่อนประมวลผล"
    elif not processed:
        steps[3] = "active"
        next_action = "ประมวลผลไฟล์หลัง Mapping พร้อมใช้งาน"
    elif not peaks_ready:
        next_action = "กำหนดช่วง Peak ก่อนส่งออก"
    elif not output_ready:
        next_action = "พร้อมส่งออกรายงาน"
    else:
        next_action = "สร้างรายงานแล้ว พร้อมดาวน์โหลดไฟล์"

    summary = [
        ("ไฟล์สำรวจ", "โหลดแล้ว" if uploaded_name else "ยังไม่ได้โหลด", "success" if uploaded_name else "neutral"),
        ("Scheme", _current_mapping_scheme(), "success" if _is_v2_scheme(_current_mapping_scheme()) else "neutral"),
        ("Mapping", "ต้องตรวจสอบ" if mapping_needs_review else ("พร้อมใช้งาน" if mapping_rows else "ยังไม่พร้อม"), "warning" if mapping_needs_review else ("success" if mapping_rows else "neutral")),
        ("ประมวลผล", "เสร็จแล้ว" if processed else "ยังไม่ได้ประมวลผล", "success" if processed else "neutral"),
        ("Peak", str(peak_state["summary_text"]), str(peak_state["summary_kind"])),
        ("ส่งออก", "สร้างแล้ว" if output_ready else ("พร้อมสร้างรายงาน" if processed and peaks_ready and excel_ready else "ยังไม่พร้อม"), "success" if output_ready or (processed and peaks_ready and excel_ready) else "neutral"),
    ]
    legacy_state = {"steps": steps, "summary": summary, "next_action": next_action}
    workflow_state = _workflow_state_for_mode(WORKFLOW_SINGLE_MODE)
    if workflow_state is not None:
        return _single_shell_from_authoritative_readiness(workflow_state.readiness, legacy_state)
    return legacy_state


def _batch_shell_from_authoritative_readiness(
    readiness: WorkflowReadiness,
    legacy_state: dict[str, object],
) -> dict[str, object]:
    """Render Batch engineering stages from the stored UX-0 state."""

    steps = ["pending"] * 6
    steps[0] = "completed" if readiness.source else "active"
    steps[1] = list(legacy_state["steps"])[1]
    steps[2] = "completed" if readiness.mapping else ("active" if readiness.source else "pending")
    steps[3] = "completed" if readiness.analysis else ("active" if readiness.mapping else "pending")
    steps[4] = "ready" if readiness.review else ("active" if readiness.analysis else "pending")
    steps[5] = "completed" if readiness.export else ("active" if readiness.review else "pending")

    if not readiness.source:
        next_action = "เริ่มจากอัปโหลดไฟล์ TMC Excel ที่แถบด้านซ้าย"
    elif not readiness.mapping:
        next_action = "เปิด Mapping Preset สำหรับ Batch ที่แถบด้านซ้าย"
    elif not readiness.analysis:
        next_action = "วิเคราะห์ Batch หลังไฟล์และ Mapping Preset พร้อม"
    elif not readiness.review:
        next_action = "กำหนดช่วง Peak ให้ครบทุกไฟล์ที่วิเคราะห์สำเร็จ"
    elif not readiness.export:
        next_action = "พร้อมสร้าง Batch ZIP"
    else:
        next_action = "สร้าง Batch ZIP แล้ว พร้อมดาวน์โหลดไฟล์"

    legacy_summary = list(legacy_state["summary"])
    summary = [
        (
            legacy_summary[0][0],
            legacy_summary[0][1] if readiness.source else "ยังไม่ได้โหลด",
            "success" if readiness.source else "neutral",
        ),
        legacy_summary[1],
        (legacy_summary[2][0], "พร้อมใช้งาน" if readiness.mapping else "ยังไม่พร้อม", "success" if readiness.mapping else "neutral"),
        (
            legacy_summary[3][0],
            "วิเคราะห์แล้ว" if readiness.analysis else "ต้องวิเคราะห์ใหม่",
            "success" if readiness.analysis else "warning",
        ),
        (
            legacy_summary[4][0],
            legacy_summary[4][1] if readiness.review else ("รอกำหนด Peak" if readiness.analysis else "ยังไม่มีไฟล์สำเร็จ"),
            "success" if readiness.review else ("warning" if readiness.analysis else "neutral"),
        ),
        (
            legacy_summary[5][0],
            "สร้าง Batch ZIP แล้ว" if readiness.export else ("รอสร้าง Batch ZIP" if readiness.review else "ยังไม่พร้อม"),
            "success" if readiness.export else "neutral",
        ),
    ]
    return {"steps": steps, "summary": summary, "next_action": next_action, "readiness": _readiness_payload(readiness)}


def derive_batch_workflow_state(
    *,
    uploaded_count: int,
    batch_mapping_ready: bool,
    batch_signature: tuple[object, ...],
) -> dict[str, object]:
    metadata_rows = st.session_state.get("tmc_batch_file_metadata_table") or []
    pce_ready = bool(_current_pce_factors_from_state())
    batch_analysis = st.session_state.get("tmc_batch_analysis_result")
    batch_result = st.session_state.get("tmc_batch_export_result")
    batch_export_stale = bool(st.session_state.get("tmc_batch_export_stale"))
    previous_signature = st.session_state.get("tmc_batch_input_signature")
    batch_stale = bool(st.session_state.get("tmc_batch_stale")) or batch_change_invalidates(
        previous_signature,
        batch_signature,
        batch_analysis is not None,
    )
    successful_items = list(batch_analysis.successful_items) if batch_analysis else []
    successful_count = len(successful_items)
    confirmed_count = sum(1 for item in successful_items if item.confirmed_AM_peak and item.confirmed_PM_peak)
    peaks_ready = bool(successful_items) and confirmed_count == successful_count

    steps = ["pending"] * 6
    if uploaded_count:
        steps[0] = "completed"
    else:
        steps[0] = "active"
    if uploaded_count and metadata_rows and pce_ready:
        steps[1] = "completed"
    elif uploaded_count:
        steps[1] = "active"
    if batch_mapping_ready:
        steps[2] = "completed"
    elif uploaded_count:
        steps[2] = "active"
    if batch_analysis and not batch_stale:
        steps[3] = "completed"
    elif batch_stale:
        steps[3] = "warning"
    elif uploaded_count and batch_mapping_ready and pce_ready:
        steps[3] = "active"
    if peaks_ready:
        steps[4] = "ready"
    elif batch_analysis and successful_items:
        steps[4] = "warning" if confirmed_count else "active"
    if batch_result and not batch_export_stale:
        steps[5] = "completed"
    elif peaks_ready and not batch_stale and batch_export_stale:
        steps[5] = "warning"
    elif peaks_ready and not batch_stale:
        steps[5] = "ready"

    if not uploaded_count:
        next_action = "เริ่มจากอัปโหลดไฟล์ TMC Excel ที่แถบด้านซ้าย"
    elif not batch_mapping_ready:
        next_action = "เปิด Mapping Preset สำหรับ Batch ที่แถบด้านซ้าย"
    elif batch_stale:
        next_action = "ข้อมูล Batch มีการเปลี่ยนแปลง กรุณาวิเคราะห์ Batch ใหม่"
    elif not batch_analysis:
        next_action = "วิเคราะห์ Batch หลังไฟล์และ Mapping Preset พร้อม"
    elif successful_items and not peaks_ready:
        next_action = "กำหนดช่วง Peak ให้ครบทุกไฟล์ที่วิเคราะห์สำเร็จ"
    elif peaks_ready and batch_export_stale:
        next_action = "ข้อมูลส่งออกมีการเปลี่ยนแปลง กรุณาสร้าง Batch ZIP ใหม่"
    elif peaks_ready and not batch_result:
        next_action = "พร้อมสร้าง Batch ZIP"
    else:
        next_action = "สร้าง Batch ZIP แล้ว พร้อมดาวน์โหลดไฟล์"

    summary = [
        ("ไฟล์สำรวจ", f"{uploaded_count:,} ไฟล์" if uploaded_count else "ยังไม่ได้โหลด", "success" if uploaded_count else "neutral"),
        ("Scheme", str(st.session_state.get("tmc_batch_mapping_code_scheme") or MOVEMENT_SCHEME_V1), "warning" if _is_v2_scheme(str(st.session_state.get("tmc_batch_mapping_code_scheme") or MOVEMENT_SCHEME_V1)) else "neutral"),
        ("Mapping Preset", "พร้อมใช้งาน" if batch_mapping_ready else "ยังไม่พร้อม", "success" if batch_mapping_ready else "neutral"),
        ("Batch Analysis", "ต้องวิเคราะห์ใหม่" if batch_stale else ("วิเคราะห์แล้ว" if batch_analysis else "ยังไม่ได้วิเคราะห์"), "warning" if batch_stale else ("success" if batch_analysis else "neutral")),
        ("Peak", f"กำหนดแล้ว {confirmed_count:,}/{successful_count:,} ไฟล์" if successful_count else "ยังไม่มีไฟล์สำเร็จ", "success" if peaks_ready else ("warning" if successful_count else "neutral")),
        ("ส่งออก", "ต้องสร้าง ZIP ใหม่" if batch_export_stale and peaks_ready and not batch_stale else ("สร้าง Batch ZIP แล้ว" if batch_result else ("พร้อมสร้าง Batch ZIP" if peaks_ready and not batch_stale else "ยังไม่พร้อม")), "warning" if batch_export_stale and peaks_ready and not batch_stale else ("success" if batch_result or (peaks_ready and not batch_stale) else "neutral")),
    ]
    legacy_state = {"steps": steps, "summary": summary, "next_action": next_action}
    workflow_state = _workflow_state_for_mode(WORKFLOW_BATCH_MODE)
    if workflow_state is not None:
        return _batch_shell_from_authoritative_readiness(workflow_state.readiness, legacy_state)
    return legacy_state


def _render_workflow_shell(
    *,
    is_single_file_mode: bool,
    uploaded_name: str | None,
    uploaded_count: int,
    batch_mapping_ready: bool,
    batch_signature: tuple[object, ...],
    export_mode: str | None,
    excel_com_status: ExcelComStatus,
) -> None:
    labels = ["อัปโหลดไฟล์", "ตั้งค่างาน", "Mapping", "ประมวลผล", "ตรวจ Peak", "ส่งออก"]
    state = (
        derive_single_workflow_state(uploaded_name, export_mode, excel_com_status)
        if is_single_file_mode
        else derive_batch_workflow_state(
            uploaded_count=uploaded_count,
            batch_mapping_ready=batch_mapping_ready,
            batch_signature=batch_signature,
        )
    )
    mode_label = "ประมวลผลไฟล์เดียว" if is_single_file_mode else "ประมวลผลหลายไฟล์"
    html = [
        '<div class="tmc-workflow-shell">',
        '<div class="tmc-workflow-shell-head">',
        f'<div class="tmc-workflow-mode">{escape(mode_label)}</div>',
        f'<div class="tmc-workflow-next">{escape(str(state["next_action"]))}</div>',
        "</div>",
        '<div class="tmc-workflow-stepper">',
    ]
    for index, (label, status) in enumerate(zip(labels, state["steps"]), start=1):
        dot = "✓" if status in {"completed", "ready"} else ("!" if status == "warning" else str(index))
        html.append(
            f'<div class="tmc-workflow-step tmc-workflow-{escape(str(status))}">'
            '<div class="tmc-workflow-step-rail">'
            f'<div class="tmc-workflow-dot">{escape(dot)}</div>'
            "</div>"
            '<div class="tmc-workflow-step-copy">'
            f'<div class="tmc-workflow-step-label">{escape(label)}</div>'
            f'<div class="tmc-workflow-step-state">{escape(_workflow_status_label(str(status)))}</div>'
            "</div>"
            "</div>"
        )
    html.append("</div>")
    html.append('<div class="tmc-workflow-summary">')
    for label, value, kind in state["summary"]:
        html.append(_workflow_summary_chip(str(label), str(value), str(kind)))
    html.append("</div></div>")
    st.markdown("".join(html), unsafe_allow_html=True)


def _probe_excel_com_for_ui(force: bool = False) -> ExcelComStatus:
    status_key = "tmc_excel_com_status"
    if force or status_key not in st.session_state:
        st.session_state[status_key] = probe_excel_com()
    return st.session_state[status_key]


def _single_export_mode_options(excel_com_status: ExcelComStatus) -> list[str]:
    return [EXCEL_TEMPLATE_EXPORT_MODE, SAFE_PNG_EXPORT_MODE] if excel_com_status.available else [SAFE_PNG_EXPORT_MODE]


def _batch_export_mode_options(excel_com_status: ExcelComStatus, movement_code_scheme: str = MOVEMENT_SCHEME_V1) -> list[str]:
    if _is_v2_scheme(movement_code_scheme):
        return [BATCH_SAFE_PNG_EXPORT_LABEL]
    return [BATCH_EXCEL_TEMPLATE_EXPORT_LABEL, BATCH_SAFE_PNG_EXPORT_LABEL] if excel_com_status.available else [BATCH_SAFE_PNG_EXPORT_LABEL]


def _default_single_export_mode(excel_com_status: ExcelComStatus) -> str:
    return EXCEL_TEMPLATE_EXPORT_MODE if excel_com_status.available else SAFE_PNG_EXPORT_MODE


def _default_batch_export_mode(excel_com_status: ExcelComStatus, movement_code_scheme: str = MOVEMENT_SCHEME_V1) -> str:
    return BATCH_SAFE_PNG_EXPORT_LABEL if _is_v2_scheme(movement_code_scheme) or not excel_com_status.available else BATCH_EXCEL_TEMPLATE_EXPORT_LABEL


def _coerce_export_mode(value: str | None, options: list[str], fallback: str) -> str:
    if value in options:
        return str(value)
    return fallback if fallback in options else options[0]


def _flash_and_rerun(message: str, kind: str = "success") -> None:
    st.session_state["tmc_flash_message"] = {"message": message, "kind": kind}
    st.rerun()


def _state_value(key: str, default: object = "") -> object:
    return st.session_state.get(key, default)


def _pce_factor_records(factors: dict[str, float] | None = None) -> list[dict[str, object]]:
    selected = normalize_pce_factors(factors)
    return [
        {"vehicle_class": vehicle_class, "pce_factor": selected[vehicle_class]}
        for vehicle_class in get_default_pce_factors()
    ]


def _ensure_pce_factor_state() -> None:
    st.session_state.setdefault("pce_editor_version", 0)
    if "pce_factors_table" not in st.session_state:
        st.session_state["pce_factors_table"] = _pce_factor_records()


def _current_pce_factor_input() -> object:
    return st.session_state.get("pce_factors_table")


def _current_pce_factors_from_state() -> dict[str, float]:
    _ensure_pce_factor_state()
    return normalize_pce_factors(_current_pce_factor_input())


def _store_selected_pce_factors(selected: dict[str, float]) -> None:
    records = _pce_factor_records(selected)
    previous = st.session_state.get("tmc_selected_pce_factors")
    if st.session_state.get("pce_factors_table") != records:
        if previous is not None and previous != selected:
            _mark_processed_outputs_stale_for_pce_change()
        st.session_state["pce_factors_table"] = records
        st.session_state["pce_editor_version"] = int(st.session_state.get("pce_editor_version", 0) or 0) + 1
    st.session_state["tmc_selected_pce_factors"] = selected


def _mark_processed_outputs_stale_for_pce_change() -> None:
    if "tmc_processed" in st.session_state:
        st.session_state["tmc_pce_results_stale"] = True
    for stale_key in [
        "tmc_output",
        "am_peak_period_select",
        "pm_peak_period_select",
        "tmc_confirmed_am_peak_start",
        "tmc_confirmed_am_peak_end",
        "tmc_confirmed_pm_peak_start",
        "tmc_confirmed_pm_peak_end",
    ]:
        st.session_state.pop(stale_key, None)


def _processed_pce_results_stale(processed: dict[str, object] | None, selected_pce_factors: dict[str, float]) -> bool:
    if not processed:
        return False
    return bool(st.session_state.get("tmc_pce_results_stale")) or not pce_factors_equal(
        processed.get("pce_factors"),
        selected_pce_factors,
    )


def _render_pce_factor_editor() -> dict[str, float]:
    _ensure_pce_factor_state()
    with st.expander("ค่าเทียบเท่ารถยนต์นั่ง (PCE)", expanded=False):
        st.caption(
            "ค่า PCE ใช้สำหรับแปลงจำนวนยานพาหนะเป็น PCU หากไม่แก้ไข โปรแกรมจะใช้ค่าเริ่มต้นตามมาตรฐานที่กำหนดไว้"
        )
        if st.button("Reset PCE เป็นค่าเริ่มต้น", key="reset_pce_factors"):
            st.session_state["pce_factors_table"] = _pce_factor_records()
            st.session_state["pce_editor_version"] = int(st.session_state.get("pce_editor_version", 0) or 0) + 1
            _mark_processed_outputs_stale_for_pce_change()
            st.success("Reset PCE เป็นค่าเริ่มต้นแล้ว")

        version = int(st.session_state.get("pce_editor_version", 0) or 0)
        edited = st.data_editor(
            pd.DataFrame(st.session_state["pce_factors_table"]),
            key=f"pce_factors_editor_{version}",
            hide_index=True,
            width="stretch",
            disabled=["vehicle_class"],
            column_config={
                "vehicle_class": st.column_config.TextColumn("vehicle_class"),
                "pce_factor": st.column_config.NumberColumn("pce_factor", format="%.3f"),
            },
        )
        validation = validate_pce_factors(edited)
        selected = validation.factors
        _store_selected_pce_factors(selected)

        for warning in validation.warnings:
            st.warning(warning)

        traceability = pce_factor_traceability_frame(selected)
        overrides = traceability[traceability["source"] == "user_override"]
        if not overrides.empty:
            override_text = ", ".join(
                f"{row.vehicle_class}={float(row.pce_factor):g}" for row in overrides.itertuples(index=False)
            )
            st.info(f"PCE user overrides: {override_text}")
    return _current_pce_factors_from_state()


def _pce_override_summary(selected_pce_factors: dict[str, float]) -> tuple[bool, str]:
    traceability = pce_factor_traceability_frame(selected_pce_factors)
    overrides = traceability[traceability["source"] == "user_override"]
    if overrides.empty:
        return False, "ใช้ค่าเริ่มต้น"
    override_text = ", ".join(
        f"{row.vehicle_class}={float(row.pce_factor):g}" for row in overrides.itertuples(index=False)
    )
    return True, override_text


def _current_mapping_for_session() -> pd.DataFrame | None:
    mapping_value = st.session_state.get("mapping_table")
    if mapping_value is None:
        return None
    try:
        return pd.DataFrame(mapping_value)
    except Exception:
        return None


def _confirmed_peaks_from_state() -> dict[str, str]:
    confirmed = dict(st.session_state.get("tmc_loaded_confirmed_peaks") or {})
    for state_key, peak_key in {
        "tmc_confirmed_am_peak_start": "am_peak_start",
        "tmc_confirmed_am_peak_end": "am_peak_end",
        "tmc_confirmed_pm_peak_start": "pm_peak_start",
        "tmc_confirmed_pm_peak_end": "pm_peak_end",
    }.items():
        if st.session_state.get(state_key):
            confirmed[peak_key] = st.session_state[state_key]
    if confirmed:
        confirmed.setdefault("peak_selection_source", PEAK_SELECTION_USER_CONFIRMED)
    return confirmed


def _workflow_peak_windows_from_setup(setup: dict[str, object]) -> dict[str, tuple[object, object]]:
    return {
        "AM": (setup.get("am_peak_window_start", AM_WINDOW[0]), setup.get("am_peak_window_end", AM_WINDOW[1])),
        "PM": (setup.get("pm_peak_window_start", PM_WINDOW[0]), setup.get("pm_peak_window_end", PM_WINDOW[1])),
    }


def _sync_single_workflow_from_state(
    *,
    source_bytes: bytes | None = None,
    source_revision: object = _WORKFLOW_SOURCE_UNSET,
    source_file_name: str | None = None,
    export_mode: str | None,
) -> WorkflowTransition:
    setup = get_current_setup_from_state(source_file_name or "")
    analysis_present = bool(st.session_state.get("tmc_processed")) and not bool(
        st.session_state.get("tmc_pce_results_stale")
    )
    revisions = _single_workflow_revisions(
        source_bytes=source_bytes,
        source_revision=source_revision,
        source_file_name=source_file_name,
        mapping=st.session_state.get("mapping_table") or [],
        pce_factors=_current_pce_factors_from_state(),
        peak_mode=str(setup.get("peak_mode") or DEFAULT_PEAK_MODE),
        peak_windows=_workflow_peak_windows_from_setup(setup),
        movement_code_scheme=_current_mapping_scheme(),
        setup=setup,
        export_mode=export_mode,
        confirmed_peaks=_confirmed_peaks_from_state(),
        analysis_present=analysis_present,
    )
    readiness = WorkflowReadiness(
        source=bool(
            source_file_name
            or source_bytes is not None
            or (source_revision is not _WORKFLOW_SOURCE_UNSET and source_revision is not None)
        ),
        mapping=bool(st.session_state.get("mapping_table")),
        analysis=analysis_present,
        review=bool(_single_effective_peak_state().get("ready")),
        export=st.session_state.get("tmc_output") is not None,
    )
    return _sync_workflow_contract(WORKFLOW_SINGLE_MODE, revisions, readiness)


def _build_session_from_state(uploaded_name: str | None, uploaded_size: int | None) -> dict[str, object]:
    setup = get_current_setup_from_state(uploaded_name or st.session_state.get("tmc_loaded_source_file_name", ""))
    peak_settings = {
        "peak_mode": setup.get("peak_mode", DEFAULT_PEAK_MODE),
        "am_peak_window_start": setup.get("am_peak_window_start", AM_WINDOW[0]),
        "am_peak_window_end": setup.get("am_peak_window_end", AM_WINDOW[1]),
        "pm_peak_window_start": setup.get("pm_peak_window_start", PM_WINDOW[0]),
        "pm_peak_window_end": setup.get("pm_peak_window_end", PM_WINDOW[1]),
        **_confirmed_peaks_from_state(),
    }
    return build_project_session(
        metadata={
            "project_name": setup.get("project_name", ""),
            "tmc_id": setup.get("tmc_id", ""),
            "tmc_title": setup.get("tmc_title", ""),
            "survey_point": setup.get("survey_point", ""),
            "survey_date_text": setup.get("survey_date_text", ""),
            "weather": setup.get("weather", ""),
            "responsible_party": setup.get("responsible_party", ""),
            "survey_period": setup.get("survey_period", ""),
        },
        directions={
            "north_label": setup.get("north_label", ""),
            "south_label": setup.get("south_label", ""),
            "east_label": setup.get("east_label", ""),
            "west_label": setup.get("west_label", ""),
            "north_road": setup.get("north_road", ""),
            "south_road": setup.get("south_road", ""),
            "east_road": setup.get("east_road", ""),
            "west_road": setup.get("west_road", ""),
            "caption_text": setup.get("caption_text", ""),
            "show_u_turn": bool(setup.get("show_u_turn", True)),
        },
        mapping=_current_mapping_for_session(),
        movement_code_scheme=_current_mapping_scheme(),
        detected_sheet_names=st.session_state.get("tmc_detected_sheet_names", []),
        pce_factors=_current_pce_factors_from_state(),
        peak_settings=peak_settings,
        export_settings={
            "use_template_report_layout": _use_template_layout_for_export(str(_state_value("report_export_mode", SAFE_PNG_EXPORT_MODE))),
            "use_excel_com_native_charts": _use_template_layout_for_export(str(_state_value("report_export_mode", SAFE_PNG_EXPORT_MODE))),
            "template_version": TEMPLATE_VERSION,
            "template_name": Path(DEFAULT_TEMPLATE_PATH).name,
            "template_path": str(DEFAULT_TEMPLATE_PATH),
            "template_map_name": Path(DEFAULT_TEMPLATE_MAP_PATH).name,
            "template_map_path": str(DEFAULT_TEMPLATE_MAP_PATH),
        },
        source_file_name=uploaded_name or st.session_state.get("tmc_loaded_source_file_name", ""),
        source_file_size=uploaded_size,
    )


def _render_project_session_section(uploaded_name: str | None, uploaded_size: int | None, *, compact: bool = False) -> None:
    st.subheader("Project Session")

    project_upload = st.file_uploader(
        "เปิด Project Session",
        type=["json"],
        key="project_session_upload",
        help="เปิดไฟล์ .tmcproj.json ระบบจะใช้ค่าหลังจากกดปุ่มใช้ค่าจาก Session ที่โหลด",
    )
    if project_upload is not None:
        project_bytes = project_upload.getvalue()
        upload_identity = (project_upload.name, hashlib.sha256(project_bytes).hexdigest())
        if st.session_state.get("tmc_project_session_upload_identity") != upload_identity:
            try:
                loaded = session_from_json(project_bytes)
            except ProjectSessionError as exc:
                st.session_state.pop("tmc_loaded_project_session", None)
                st.session_state["tmc_project_session_error"] = str(exc)
            else:
                st.session_state["tmc_loaded_project_session"] = loaded.session
                st.session_state["tmc_loaded_project_session_warnings"] = list(loaded.warnings)
                st.session_state["tmc_loaded_source_file_name"] = loaded.session.get("source_file_name", "")
                st.session_state.pop("tmc_project_session_error", None)
            st.session_state["tmc_project_session_upload_identity"] = upload_identity

    if st.session_state.get("tmc_project_session_error"):
        st.error(st.session_state["tmc_project_session_error"])

    loaded_session = st.session_state.get("tmc_loaded_project_session")
    if loaded_session:
        for warning in st.session_state.get("tmc_loaded_project_session_warnings", []):
            st.warning(warning)
        mismatch = source_file_mismatch_warning(loaded_session, uploaded_name)
        if mismatch:
            st.warning(mismatch)
            st.info("กรุณาอัปโหลดไฟล์สำรวจเดิมหรือไฟล์ที่ต้องการใช้กับ Session นี้ แล้วกดประมวลผลใหม่")
        metadata = loaded_session.get("metadata", {})
        mapping_rows = loaded_session.get("mapping", {}).get("rows", [])
        peaks = loaded_session.get("peaks", {})
        preview = {
            "ไฟล์ต้นทาง": loaded_session.get("source_file_name", ""),
            "โครงการ": metadata.get("project_name", "") or metadata.get("tmc_title", ""),
            "จำนวนแถว Mapping": len(mapping_rows),
            "AM Peak": f"{peaks.get('am_peak_start', '')}-{peaks.get('am_peak_end', '')}".strip("-"),
            "PM Peak": f"{peaks.get('pm_peak_start', '')}-{peaks.get('pm_peak_end', '')}".strip("-"),
        }
        with st.expander("ตัวอย่างข้อมูล Session ที่โหลด", expanded=not compact):
            st.write(preview)
        if st.button("ใช้ค่าจาก Session ที่โหลด", key="apply_project_session"):
            changed = apply_session_to_state(loaded_session, st.session_state)
            update_setup_state_from_widgets()
            _hydrate_setup_widgets_from_state()
            previous_workflow_state = _workflow_state_for_mode(WORKFLOW_SINGLE_MODE)
            _sync_single_workflow_from_state(
                source_revision=(
                    previous_workflow_state.revisions.source
                    if previous_workflow_state is not None
                    else _WORKFLOW_SOURCE_UNSET
                ),
                source_file_name=uploaded_name or st.session_state.get("tmc_loaded_source_file_name", ""),
                export_mode=str(st.session_state.get("report_export_mode") or SAFE_PNG_EXPORT_MODE),
            )
            st.success(f"ใช้ค่า Session แล้ว อัปเดตค่าปัจจุบัน {len(changed)} รายการ")
            if changed and not compact:
                st.info("ค่าตั้งค่างานปัจจุบันถูกแทนที่ด้วยข้อมูลจาก Project Session")
    else:
        st.caption("เปิด Project Session เพื่อดูตัวอย่างก่อนใช้ค่า")
        st.button("ใช้ค่าจาก Session ที่โหลด", key="apply_project_session_disabled", disabled=True)

    session = _build_session_from_state(uploaded_name, uploaded_size)
    session_bytes = session_to_json_bytes(session)
    st.session_state["tmc_project_session_bytes"] = session_bytes
    filename_seed = st.session_state.get("tmc_id_input") or st.session_state.get("tmc_title_input") or uploaded_name
    session_filename = safe_project_session_filename(filename_seed)
    st.session_state["tmc_project_session_filename"] = session_filename
    st.download_button(
        "ดาวน์โหลด Project Session",
        data=download_buffer(session_bytes),
        file_name=session_filename,
        mime=PROJECT_SESSION_MIME,
        key="download_project_session_sidebar" if compact else "download_project_session_main",
        help="บันทึกเฉพาะค่าตั้งค่า ไม่รวมไฟล์ Excel สำรวจต้นฉบับ",
    )
    if not compact:
        st.caption("Project Session บันทึกเฉพาะค่าตั้งค่าและข้อมูลไฟล์ต้นทาง ไม่รวมเนื้อหา Excel ต้นฉบับ")


def _ordered_mapping_frame(mapping: pd.DataFrame) -> pd.DataFrame:
    preferred = [
        "raw_sheet",
        "raw_direction",
        "source_stream",
        "raw_movement_label",
        "approach_direction",
        "movement_type",
        "movement_code",
        "from_leg",
        "to_leg",
        "turn_type",
        "include_in_report",
        "include_in_peak",
        "aggregation_method",
        "facility_type",
        "note",
        "remark",
    ]
    ordered = [column for column in preferred if column in mapping.columns]
    ordered.extend(column for column in mapping.columns if column not in ordered)
    return mapping[ordered].copy()


def _mapping_editor_frame(mapping: pd.DataFrame, view_mode: str, movement_code_scheme: str = MOVEMENT_SCHEME_V1) -> pd.DataFrame:
    ordered = _ordered_mapping_frame(normalize_mapping_for_scheme(mapping, movement_code_scheme))
    if _is_v2_scheme(movement_code_scheme):
        for column in ("approach_direction", "movement_type"):
            if column not in ordered.columns:
                ordered[column] = ""
    if view_mode != "Basic":
        return ordered
    basic_columns = [
        "raw_sheet",
        "raw_direction",
        "source_stream",
        "movement_code",
        "raw_movement_label",
        "include_in_report",
        "include_in_peak",
    ]
    visible = [column for column in basic_columns if column in ordered.columns]
    return ordered[visible].copy()


_MOVEMENT_LABELS_TH = {
    "L": "ซ้าย",
    "T": "ตรง",
    "R": "ขวา",
    "U": "กลับรถ",
}


def _movement_label_th(turn_type: object) -> str:
    return _MOVEMENT_LABELS_TH.get(str(turn_type or "").strip(), "")


def _basic_mapping_derived_summary(mapping: pd.DataFrame, movement_code_scheme: str = MOVEMENT_SCHEME_V1) -> pd.DataFrame:
    normalized = normalize_mapping_for_scheme(mapping, movement_code_scheme)
    valid_codes = set(APPROACH_MOVEMENT_CODES if _is_v2_scheme(movement_code_scheme) else MOVEMENT_CODE_OPTIONS)
    rows = []
    for _, row in normalized.iterrows():
        code = str(row.get("movement_code") or "").strip()
        from_leg = str(row.get("from_leg") or "").strip()
        to_leg = str(row.get("to_leg") or "").strip()
        turn_type = str(row.get("movement_type" if _is_v2_scheme(movement_code_scheme) else "turn_type") or "").strip()
        if not code:
            status = "Missing movement_code"
        elif code not in valid_codes:
            status = "รหัสทิศทางไม่ถูกต้อง: ใช้รูปแบบเช่น NE, NS, NW, NU"
        else:
            status = "Valid"
        rows.append(
            {
                "Sheet": row.get("raw_sheet", ""),
                "Code": code,
                "Derived movement": f"{from_leg} → {to_leg}" if from_leg and to_leg else "",
                "Type": turn_type,
                "Thai label": _movement_label_th(turn_type),
                "Status": status,
            }
        )
    return pd.DataFrame(rows, columns=["Sheet", "Code", "Derived movement", "Type", "Thai label", "Status"])


def _mapping_editor_column_config(movement_code_scheme: str, movement_code_options: list[str]) -> dict[str, object]:
    editor_labels = _mapping_editor_labels(movement_code_scheme)
    derived_v2 = _is_v2_scheme(movement_code_scheme)
    derived_fields = True
    return {
        "raw_sheet": st.column_config.TextColumn(editor_labels["raw_sheet"], disabled=True),
        "raw_direction": st.column_config.TextColumn(editor_labels["raw_direction"], disabled=True),
        "source_stream": st.column_config.SelectboxColumn(
            editor_labels["source_stream"],
            options=SOURCE_STREAM_OPTIONS,
            required=False,
            default="mainline",
        ),
        "raw_movement_label": st.column_config.TextColumn(editor_labels["raw_movement_label"], required=False),
        "movement_code": st.column_config.SelectboxColumn(
            editor_labels["movement_code"],
            options=movement_code_options,
            required=True,
            help=_movement_code_help_text(movement_code_scheme),
        ),
        "approach_direction": st.column_config.SelectboxColumn(
            editor_labels.get("approach_direction", "approach_direction"),
            options=["", "N", "S", "E", "W"],
            required=False,
            disabled=derived_v2,
        ),
        "movement_type": st.column_config.SelectboxColumn(
            editor_labels.get("movement_type", "movement_type"),
            options=["", "L", "T", "R", "U"],
            required=False,
            disabled=derived_v2,
        ),
        "from_leg": st.column_config.SelectboxColumn(
            "from_leg",
            options=["", *LEG_OPTIONS],
            required=not derived_fields,
            disabled=derived_fields,
        ),
        "to_leg": st.column_config.SelectboxColumn(
            "to_leg",
            options=["", *LEG_OPTIONS],
            required=not derived_fields,
            disabled=derived_fields,
        ),
        "turn_type": st.column_config.SelectboxColumn(
            "turn_type",
            options=["", *TURN_TYPE_OPTIONS, "L", "T", "R", "U"],
            required=False,
            disabled=derived_fields,
        ),
        "facility_type": st.column_config.SelectboxColumn(
            "facility_type",
            options=["", *FACILITY_TYPE_OPTIONS],
            required=True,
            default="at_grade",
        ),
        "include_in_peak": st.column_config.CheckboxColumn(editor_labels["include_in_peak"], default=True),
        "include_in_report": st.column_config.CheckboxColumn(editor_labels["include_in_report"], default=True),
        "aggregation_method": st.column_config.SelectboxColumn(
            "aggregation_method",
            options=AGGREGATION_METHOD_OPTIONS,
            required=False,
            default="sum",
        ),
    }


def _mapping_editor_key(view_mode: str, version: int) -> str:
    return f"mapping_editor_{view_mode.lower()}_{version}"


def _basic_mapping_widget_key(column: str, version: int, row_index: int) -> str:
    return f"mapping_basic_{column}_{version}_{row_index}"


def _option_index(options: list[str], value: object) -> int:
    text = str(value or "").strip()
    return options.index(text) if text in options else 0


def _apply_basic_mapping_widget_state(
    editor_frame: pd.DataFrame,
    version: int,
    state: MutableMapping[str, object] | None = None,
) -> pd.DataFrame:
    source_state = st.session_state if state is None else state
    edited = pd.DataFrame(editor_frame).reset_index(drop=True).copy()
    for row_index in range(len(edited)):
        for column in ("source_stream", "movement_code", "raw_movement_label", "include_in_report", "include_in_peak"):
            key = _basic_mapping_widget_key(column, version, row_index)
            if key in source_state and column in edited.columns:
                edited.at[row_index, column] = source_state[key]
    return edited


def _basic_mapping_widget_state_exists(version: int, state: MutableMapping[str, object]) -> bool:
    prefix = f"mapping_basic_"
    suffix = f"_{version}_"
    return any(str(key).startswith(prefix) and suffix in str(key) for key in state)


def _apply_mapping_editor_widget_state(editor_frame: pd.DataFrame, editor_state: object) -> pd.DataFrame:
    edited = pd.DataFrame(editor_frame).reset_index(drop=True).copy()
    if not isinstance(editor_state, dict):
        return edited

    for row_index in sorted((int(row) for row in editor_state.get("deleted_rows", []) if str(row).isdigit()), reverse=True):
        if 0 <= row_index < len(edited):
            edited = edited.drop(index=row_index).reset_index(drop=True)

    for row in editor_state.get("added_rows", []):
        if isinstance(row, dict):
            edited = pd.concat([edited, pd.DataFrame([row])], ignore_index=True)

    for row_key, changes in (editor_state.get("edited_rows") or {}).items():
        if not str(row_key).isdigit() or not isinstance(changes, dict):
            continue
        row_index = int(row_key)
        if 0 <= row_index < len(edited):
            for column, value in changes.items():
                if column in edited.columns:
                    edited.at[row_index, column] = value

    for cell_key, value in (editor_state.get("edited_cells") or {}).items():
        row_text, _, column = str(cell_key).partition(":")
        if row_text.isdigit() and column in edited.columns:
            row_index = int(row_text)
            if 0 <= row_index < len(edited):
                edited.at[row_index, column] = value

    return edited


def _render_basic_mapping_editor(
    editor_frame: pd.DataFrame,
    movement_code_options: list[str],
    version: int,
) -> pd.DataFrame:
    edited = pd.DataFrame(editor_frame).reset_index(drop=True).copy()
    headers = st.columns([1.2, 1.2, 1.1, 1.2, 1.2, 0.8, 0.8])
    for column, label in zip(
        headers,
        ["Sheet", "Raw", "Source", "Movement", "Label", "Report", "Peak"],
        strict=True,
    ):
        column.caption(label)

    for row_index, row in edited.iterrows():
        cols = st.columns([1.2, 1.2, 1.1, 1.2, 1.2, 0.8, 0.8])
        cols[0].write(str(row.get("raw_sheet", "")))
        cols[1].write(str(row.get("raw_direction", "")))
        edited.at[row_index, "source_stream"] = cols[2].selectbox(
            "source_stream",
            options=SOURCE_STREAM_OPTIONS,
            index=_option_index(SOURCE_STREAM_OPTIONS, row.get("source_stream")),
            key=_basic_mapping_widget_key("source_stream", version, row_index),
            label_visibility="collapsed",
        )
        edited.at[row_index, "movement_code"] = cols[3].selectbox(
            "movement_code",
            options=movement_code_options,
            index=_option_index(movement_code_options, row.get("movement_code")),
            key=_basic_mapping_widget_key("movement_code", version, row_index),
            label_visibility="collapsed",
        )
        edited.at[row_index, "raw_movement_label"] = cols[4].text_input(
            "raw_movement_label",
            value=str(row.get("raw_movement_label", "")),
            key=_basic_mapping_widget_key("raw_movement_label", version, row_index),
            label_visibility="collapsed",
        )
        edited.at[row_index, "include_in_report"] = cols[5].checkbox(
            "include_in_report",
            value=bool(row.get("include_in_report", True)),
            key=_basic_mapping_widget_key("include_in_report", version, row_index),
            label_visibility="collapsed",
        )
        edited.at[row_index, "include_in_peak"] = cols[6].checkbox(
            "include_in_peak",
            value=bool(row.get("include_in_peak", True)),
            key=_basic_mapping_widget_key("include_in_peak", version, row_index),
            label_visibility="collapsed",
        )
    return edited


def _mapping_from_editor_widget_state(
    base_mapping: pd.DataFrame,
    view_mode: str,
    movement_code_scheme: str,
    editor_key: str,
    state: MutableMapping[str, object] | None = None,
    apply_basic_widgets: bool = True,
) -> pd.DataFrame:
    source_state = st.session_state if state is None else state
    version = int(editor_key.rsplit("_", 1)[-1] or 0)
    if apply_basic_widgets and _basic_mapping_widget_state_exists(version, source_state):
        basic_visible = _mapping_editor_frame(base_mapping, "Basic", movement_code_scheme)
        basic_edited = _apply_basic_mapping_widget_state(basic_visible, version, source_state)
        base_mapping = _merge_mapping_editor_result(base_mapping, basic_edited, movement_code_scheme)

    editor_state = source_state.get(editor_key)
    if not isinstance(editor_state, dict):
        return base_mapping
    visible = _mapping_editor_frame(base_mapping, view_mode, movement_code_scheme)
    edited_visible = _apply_mapping_editor_widget_state(visible, editor_state)
    return _merge_mapping_editor_result(base_mapping, edited_visible, movement_code_scheme)


def _merge_mapping_editor_result(
    base_mapping: pd.DataFrame,
    edited_visible: pd.DataFrame,
    movement_code_scheme: str = MOVEMENT_SCHEME_V1,
) -> pd.DataFrame:
    edited = pd.DataFrame(edited_visible)
    base = _ordered_mapping_frame(base_mapping).reset_index(drop=True)
    if edited.empty and len(edited.columns) == 0:
        return base
    merged = base.reindex(range(len(edited))).copy()
    for column in edited.columns:
        merged[column] = edited[column].values
    defaults = {
        "facility_type": "at_grade",
        "aggregation_method": "sum",
        "source_stream": "mainline",
        "include_in_report": True,
        "include_in_peak": True,
    }
    for column, default in defaults.items():
        if column in merged.columns:
            merged[column] = merged[column].fillna(default)
    return _ordered_mapping_frame(normalize_mapping_for_scheme(merged, movement_code_scheme))


def _interval_total_pcu(hourly_movement: pd.DataFrame, label: str) -> str:
    if hourly_movement.empty or not label or "Total" not in hourly_movement.columns:
        return ""
    time_column = str(hourly_movement.columns[0])
    rows = hourly_interval_rows(hourly_movement)
    matched = rows[rows[time_column].astype(str).str.replace(".", ":", regex=False) == label.replace(".", ":")]
    if matched.empty:
        return ""
    value = pd.to_numeric(matched.iloc[0].get("Total"), errors="coerce")
    return format_pcu(value) if pd.notna(value) else ""


def _render_hourly_pcu_line_chart(hourly_movement: pd.DataFrame, am_peak_label: str = "", pm_peak_label: str = "") -> None:
    chart_frame = hourly_interval_rows(hourly_movement)
    if chart_frame.empty:
        _render_empty_state("No hourly PCU data", "Analyze the workbook before reviewing the hourly PCU chart.")
        return
    time_column = chart_frame.columns[0]
    if "Total" in chart_frame.columns:
        total = chart_frame[["Total"]].copy()
    else:
        value_columns = [column for column in chart_frame.columns if column != time_column]
        total = pd.DataFrame({"Total": chart_frame[value_columns].apply(pd.to_numeric, errors="coerce").sum(axis=1) if value_columns else 0})
    display = pd.DataFrame(
        {
            "เวลา": chart_frame[time_column].astype(str).map(format_time_range),
            "PCU": pd.to_numeric(total["Total"], errors="coerce").fillna(0),
        }
    )
    time_order = display["เวลา"].tolist()
    base = alt.Chart(display).encode(
        x=alt.X("เวลา:N", axis=alt.Axis(title="เวลา", labelAngle=-35), sort=time_order),
        tooltip=[
            alt.Tooltip("เวลา:N", title="ช่วงเวลา"),
            alt.Tooltip("PCU:Q", title="PCU/ชม.", format=",.0f"),
        ],
    )
    area = base.mark_area(color=CHART_PRIMARY_COLOR, opacity=0.11).encode(y=alt.Y("PCU:Q"))
    line = base.mark_line(color=CHART_PRIMARY_COLOR, strokeWidth=2.6).encode(
        y=alt.Y(
            "PCU:Q",
            axis=alt.Axis(title="PCU/ชม.", grid=True, gridColor=CHART_GRID_COLOR),
            scale=alt.Scale(zero=True),
        )
    )
    points = base.mark_point(color=CHART_PRIMARY_COLOR, filled=True, size=48).encode(y=alt.Y("PCU:Q"))
    layers = [area, line, points]
    highlights = []
    for label, period, color in (
        ("AM Peak", am_peak_label, CHART_PRIMARY_COLOR),
        ("PM Peak", pm_peak_label, CHART_PM_COLOR),
    ):
        formatted = format_time_range(period)
        matched = display[display["เวลา"].eq(formatted)]
        if not matched.empty:
            row = matched.iloc[0]
            highlights.append({"PCU": float(row["PCU"]), "ช่วง": label, "เวลา": formatted, "color": color})
    if highlights:
        highlight_frame = pd.DataFrame(highlights)
        highlight_base = alt.Chart(highlight_frame).encode(
            x=alt.X("เวลา:N", sort=time_order),
            y=alt.Y("PCU:Q"),
            color=alt.Color("color:N", scale=None, legend=None),
            tooltip=[
                alt.Tooltip("ช่วง:N", title="Peak"),
                alt.Tooltip("เวลา:N", title="ช่วงเวลา"),
                alt.Tooltip("PCU:Q", title="PCU/ชม.", format=",.0f"),
            ],
        )
        layers.append(highlight_base.mark_rule(strokeWidth=1.8, opacity=0.55))
        layers.append(highlight_base.mark_point(filled=True, size=110, stroke="white", strokeWidth=1.2))
    chart = (
        alt.layer(*layers)
        .properties(height=320)
        .configure_view(stroke=None)
        .configure_axis(labelColor=CHART_TEXT_COLOR, titleColor=CHART_TEXT_COLOR, domainColor=CHART_GRID_COLOR, tickColor=CHART_GRID_COLOR)
    )
    st.altair_chart(chart, width="stretch")


def _batch_upload_signature(files: list[object] | tuple[object, ...] | None) -> tuple[tuple[str, int, str], ...]:
    signature = []
    for file in files or []:
        data = file.getvalue()
        signature.append((Path(file.name).name, len(data), hashlib.sha256(data).hexdigest()))
    return tuple(signature)


def _metadata_signature(rows: list[dict[str, object]]) -> tuple[tuple[str, str, str, str], ...]:
    return tuple(
        (
            str(row.get("file_name", "")),
            str(row.get("survey_date_text", "")),
            safe_output_stem(str(row.get("output_stem", ""))),
            str(row.get("notes", "")),
        )
        for row in rows
    )


def _batch_peak_settings_signature(
    peak_mode: str | None,
    peak_windows: dict[str, tuple[str, str]] | None,
) -> tuple[object, ...]:
    return (
        str(peak_mode or ""),
        tuple(sorted((str(key), tuple(value)) for key, value in (peak_windows or {}).items())),
    )


def _batch_analysis_signature(
    *,
    uploads_signature: tuple[object, ...],
    preset_signature: object,
    pce_factors: dict[str, float] | None,
    peak_mode: str | None,
    peak_windows: dict[str, tuple[str, str]] | None,
) -> tuple[object, ...]:
    return (
        uploads_signature,
        preset_signature,
        tuple(sorted((pce_factors or {}).items())),
        _batch_peak_settings_signature(peak_mode, peak_windows),
    )


def _batch_export_signature(
    *,
    metadata_rows: list[dict[str, object]],
    shared_setup: dict[str, object] | None = None,
    export_mode: str | None,
    confirmed_peaks: dict[str, dict[str, str]] | None,
) -> tuple[object, ...]:
    setup = shared_setup or {}
    return (
        _metadata_signature(metadata_rows),
        tuple(
            (field, str(setup.get(field, "")))
            for field in (
                "project_name",
                "tmc_id",
                "tmc_title",
                "survey_point",
                "weather",
                "responsible_party",
                "survey_period",
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
        ),
        str(export_mode or ""),
        tuple(
            sorted(
                (
                    str(folder),
                    str(values.get("AM", "")),
                    str(values.get("PM", "")),
                )
                for folder, values in (confirmed_peaks or {}).items()
            )
        ),
    )


def _stable_batch_confirmed_peaks(
    confirmed_peaks: dict[str, dict[str, str]] | None,
    batch_analysis: object | None,
) -> dict[str, dict[str, str]]:
    """Key Batch review decisions by source file, not export folder names."""

    confirmed = confirmed_peaks or {}
    items = list(getattr(batch_analysis, "items", []) or []) if batch_analysis else []
    if not items:
        return confirmed
    return {
        str(item.file_name): dict(confirmed.get(str(item.folder_name), {}))
        for item in items
        if getattr(item, "file_name", None)
    }


def _sync_batch_workflow_from_state(
    *,
    batch_uploads: list[object] | tuple[object, ...] | None,
    mapping_preset: dict[str, object] | None,
    movement_code_scheme: str,
    metadata_rows: list[dict[str, object]],
    export_mode: str | None,
) -> WorkflowTransition:
    setup = get_current_setup_from_state("")
    batch_analysis = st.session_state.get("tmc_batch_analysis_result")
    analysis_present = batch_analysis is not None and not bool(st.session_state.get("tmc_batch_stale"))
    revisions = _batch_workflow_revisions(
        uploads=batch_uploads,
        mapping_preset=mapping_preset,
        pce_factors=_current_pce_factors_from_state(),
        peak_mode=str(setup.get("peak_mode") or DEFAULT_PEAK_MODE),
        peak_windows=_workflow_peak_windows_from_setup(setup),
        movement_code_scheme=movement_code_scheme,
        metadata_rows=metadata_rows,
        setup=setup,
        export_mode=export_mode,
        confirmed_peaks=_stable_batch_confirmed_peaks(
            st.session_state.get("tmc_batch_confirmed_peaks") or {},
            batch_analysis,
        ),
        analysis_present=analysis_present,
    )
    successful_items = list(getattr(batch_analysis, "successful_items", []) or []) if batch_analysis else []
    review_ready = bool(successful_items) and all(
        item.confirmed_AM_peak and item.confirmed_PM_peak for item in successful_items
    )
    readiness = WorkflowReadiness(
        source=bool(batch_uploads),
        mapping=mapping_preset is not None,
        analysis=analysis_present,
        review=review_ready,
        export=st.session_state.get("tmc_batch_export_result") is not None
        and not bool(st.session_state.get("tmc_batch_export_stale")),
    )
    return _sync_workflow_contract(WORKFLOW_BATCH_MODE, revisions, readiness)


def _sync_workflow_after_pce_editor(
    *,
    is_single_file_mode: bool,
    export_mode: str | None,
    source_bytes: bytes | None = None,
    source_file_name: str | None = None,
    batch_uploads: list[object] | tuple[object, ...] | None = None,
    mapping_preset: dict[str, object] | None = None,
    movement_code_scheme: str = MOVEMENT_SCHEME_V1,
    metadata_rows: list[dict[str, object]] | None = None,
) -> WorkflowTransition:
    """Re-sync authoritative readiness after the PCE editor mutates widget state."""

    if is_single_file_mode:
        return _sync_single_workflow_from_state(
            source_bytes=source_bytes,
            source_file_name=source_file_name,
            export_mode=export_mode,
        )
    return _sync_batch_workflow_from_state(
        batch_uploads=batch_uploads,
        mapping_preset=mapping_preset,
        movement_code_scheme=movement_code_scheme,
        metadata_rows=metadata_rows or [],
        export_mode=export_mode,
    )


def _mark_batch_export_stale_now() -> None:
    if st.session_state.get("tmc_batch_export_result") is not None:
        st.session_state["tmc_batch_export_stale"] = True
    st.session_state.pop("tmc_batch_export_result", None)


def _mark_batch_export_stale_if_inputs_changed(signature: tuple[object, ...]) -> bool:
    previous = st.session_state.get("tmc_batch_export_signature")
    st.session_state["tmc_batch_current_export_signature"] = signature
    if previous is not None and previous != signature:
        _mark_batch_export_stale_now()
    return bool(st.session_state.get("tmc_batch_export_stale"))


def _sync_batch_analysis_metadata_from_state() -> None:
    batch_analysis = st.session_state.get("tmc_batch_analysis_result")
    if batch_analysis is None:
        return

    metadata_rows = st.session_state.get("tmc_batch_file_metadata_table") or []
    metadata_by_name = {str(row.get("file_name", "")): row for row in metadata_rows}
    analysis_items = list(getattr(batch_analysis, "items", []) or [])
    batch_items = []
    for item in analysis_items:
        metadata = metadata_by_name.get(str(item.file_name), {})
        batch_items.append(
            BatchItem(
                file_name=item.file_name,
                workbook_bytes=getattr(item, "workbook_bytes", b""),
                survey_date_text=str(metadata.get("survey_date_text", item.survey_date_text) or ""),
                output_stem=safe_output_stem(str(metadata.get("output_stem", item.output_stem) or item.output_stem)),
                notes=str(metadata.get("notes", item.notes) or ""),
            )
        )

    output_stems = unique_safe_output_stems(batch_items)
    old_to_new_folders: dict[str, str] = {}
    for index, (item, batch_item) in enumerate(zip(analysis_items, batch_items), start=1):
        old_folder = str(item.folder_name)
        output_stem = output_stems[index - 1]
        item.survey_date_text = batch_item.survey_date_text
        item.output_stem = output_stem
        item.folder_name = batch_folder_name(index, output_stem)
        item.notes = batch_item.notes or item.notes
        old_to_new_folders[old_folder] = item.folder_name

    confirmed = st.session_state.get("tmc_batch_confirmed_peaks") or {}
    remapped_confirmed = {}
    for old_folder, values in confirmed.items():
        remapped_confirmed[old_to_new_folders.get(str(old_folder), str(old_folder))] = values
    st.session_state["tmc_batch_confirmed_peaks"] = remapped_confirmed
    selected_review_file = str(st.session_state.get("tmc_batch_selected_review_file") or "")
    if selected_review_file in old_to_new_folders:
        st.session_state["tmc_batch_selected_review_file"] = old_to_new_folders[selected_review_file]


def _sync_batch_metadata_state(batch_uploads: list[object] | tuple[object, ...] | None) -> list[dict[str, str]]:
    file_names = [Path(file.name).name for file in batch_uploads or []]
    upload_names = tuple(file_names)
    existing_rows = st.session_state.get("tmc_batch_file_metadata_table") or []
    existing_by_name = {str(row.get("file_name", "")): dict(row) for row in existing_rows}
    rows = []
    for default_row in batch_file_metadata_defaults(file_names):
        existing = existing_by_name.get(default_row["file_name"], {})
        rows.append(
            {
                "file_name": default_row["file_name"],
                "survey_date_text": str(existing.get("survey_date_text", default_row["survey_date_text"]) or ""),
                "output_stem": safe_output_stem(str(existing.get("output_stem", default_row["output_stem"]) or default_row["output_stem"])),
                "notes": str(existing.get("notes", default_row["notes"]) or ""),
            }
        )
    if st.session_state.get("tmc_batch_metadata_upload_names") != upload_names:
        st.session_state["tmc_batch_metadata_upload_names"] = upload_names
        st.session_state["tmc_batch_file_metadata_table"] = rows
        st.session_state["tmc_batch_file_metadata_editor_version"] = int(st.session_state.get("tmc_batch_file_metadata_editor_version", 0) or 0) + 1
    elif existing_rows != rows:
        st.session_state["tmc_batch_file_metadata_table"] = rows
    return st.session_state.get("tmc_batch_file_metadata_table", [])


def _batch_items_from_uploads(batch_uploads: list[object] | tuple[object, ...] | None) -> list[BatchItem]:
    metadata_rows = st.session_state.get("tmc_batch_file_metadata_table") or []
    metadata_by_name = {str(row.get("file_name", "")): row for row in metadata_rows}
    items = []
    for file in batch_uploads or []:
        file_name = Path(file.name).name
        metadata = metadata_by_name.get(file_name, {})
        items.append(
            BatchItem(
                file_name=file_name,
                workbook_bytes=file.getvalue(),
                survey_date_text=str(metadata.get("survey_date_text", "") or ""),
                output_stem=safe_output_stem(str(metadata.get("output_stem", "") or file_name)),
                notes=str(metadata.get("notes", "") or ""),
            )
        )
    return items


def _mark_batch_stale_if_inputs_changed(signature: tuple[object, ...]) -> bool:
    previous = st.session_state.get("tmc_batch_input_signature")
    st.session_state["tmc_batch_current_input_signature"] = signature
    if batch_change_invalidates(previous, signature, st.session_state.get("tmc_batch_analysis_result") is not None):
        st.session_state["tmc_batch_stale"] = True
        st.session_state["tmc_batch_export_stale"] = True
        st.session_state.pop("tmc_batch_export_result", None)
    return bool(st.session_state.get("tmc_batch_stale"))


def _mark_batch_stale_now() -> None:
    if st.session_state.get("tmc_batch_analysis_result") is not None:
        st.session_state["tmc_batch_stale"] = True
        st.session_state["tmc_batch_export_stale"] = True
        st.session_state.pop("tmc_batch_export_result", None)


def _batch_status_frame(batch_analysis, batch_result=None) -> pd.DataFrame:
    export_by_folder = {}
    if batch_result:
        export_by_folder = {row.folder_name: row for row in batch_result.summary_rows}
    rows = []
    if batch_analysis:
        for item in batch_analysis.items:
            export_row = export_by_folder.get(item.folder_name)
            rows.append(
                {
                    "file_name": item.file_name,
                    "survey_date_text": item.survey_date_text,
                    "output_stem": item.output_stem,
                    "movement_code_scheme": getattr(item, "movement_code_scheme", MOVEMENT_SCHEME_V1),
                    "status": item.status,
                    "mapping_status": item.mapping_status,
                    "AM suggested": item.suggested_AM_peak,
                    "PM suggested": item.suggested_PM_peak,
                    "AM confirmed": item.confirmed_AM_peak,
                    "PM confirmed": item.confirmed_PM_peak,
                    "total PCU": item.total_PCU,
                    "QC errors": item.QC_errors,
                    "QC warnings": item.QC_warnings,
                    "QC info": item.QC_info,
                    "export_mode_used": export_row.export_mode_used if export_row else "",
                    "export_status": export_row.export_status if export_row else "",
                    "notes": export_row.notes if export_row else item.notes,
                }
            )
    elif batch_result:
        rows = [
            {
                "file_name": row.file_name,
                "survey_date_text": row.survey_date_text,
                "output_stem": row.output_stem,
                "movement_code_scheme": getattr(row, "movement_code_scheme", MOVEMENT_SCHEME_V1),
                "status": row.status,
                "mapping_status": "",
                "AM suggested": row.suggested_AM_peak,
                "PM suggested": row.suggested_PM_peak,
                "AM confirmed": row.confirmed_AM_peak,
                "PM confirmed": row.confirmed_PM_peak,
                "total PCU": row.total_PCU,
                "QC errors": row.QC_errors,
                "QC warnings": row.QC_warnings,
                "QC info": row.QC_info,
                "export_mode_used": row.export_mode_used,
                "export_status": row.export_status,
                "notes": row.notes,
            }
            for row in batch_result.summary_rows
        ]
    return pd.DataFrame(rows)


def _batch_status_display_frame(batch_analysis, batch_result=None) -> pd.DataFrame:
    display = _batch_status_frame(batch_analysis, batch_result)
    if not display.empty:
        for column in ("status", "mapping_status", "export_status"):
            if column in display.columns:
                display[column] = display[column].map(_display_status_label)
        display = _format_display_columns(display)
    return display.rename(
        columns={
            "file_name": "ชื่อไฟล์",
            "survey_date_text": "วันที่สำรวจ",
            "output_stem": "ชื่อส่งออก",
            "status": "สถานะ",
            "mapping_status": "สถานะ Mapping",
            "AM suggested": "AM แนะนำ",
            "PM suggested": "PM แนะนำ",
            "AM confirmed": "AM กำหนดแล้ว",
            "PM confirmed": "PM กำหนดแล้ว",
            "total PCU": "PCU รวม",
            "QC errors": "QC ผิดพลาด",
            "QC warnings": "QC เตือน",
            "QC info": "QC ข้อมูล",
            "export_mode_used": "โหมดส่งออกที่ใช้",
            "export_status": "สถานะส่งออก",
            "notes": "หมายเหตุ",
        }
    )


def _batch_peak_review_display_frame(batch_analysis) -> pd.DataFrame:
    display = _batch_status_frame(batch_analysis)
    if display.empty:
        return display
    columns = [
        "file_name",
        "survey_date_text",
        "status",
        "AM suggested",
        "PM suggested",
        "AM confirmed",
        "PM confirmed",
        "QC errors",
        "QC warnings",
        "QC info",
        "notes",
    ]
    display = display[[column for column in columns if column in display.columns]]
    if "status" in display.columns:
        display["status"] = display["status"].map(_display_status_label)
    display = _format_display_columns(display)
    return display.rename(
        columns={
            "file_name": "ชื่อไฟล์",
            "survey_date_text": "วันที่สำรวจ",
            "status": "สถานะ",
            "AM suggested": "AM แนะนำ",
            "PM suggested": "PM แนะนำ",
            "AM confirmed": "AM กำหนดแล้ว",
            "PM confirmed": "PM กำหนดแล้ว",
            "QC errors": "QC ผิดพลาด",
            "QC warnings": "QC เตือน",
            "QC info": "QC ข้อมูล",
            "notes": "หมายเหตุ",
        }
    )


def _batch_qc_rows_for_ui(batch_analysis, batch_result=None) -> pd.DataFrame:
    if batch_result:
        return batch_qc_frame(batch_result.qc_rows)
    return batch_qc_frame(batch_analysis_qc_rows(batch_analysis))


def _display_status_label(value: object) -> str:
    text = str(value or "").strip()
    labels = {
        "success": "สำเร็จ",
        "failed": "ไม่สำเร็จ",
        "error": "ผิดพลาด",
        "warning": "เตือน",
        "info": "ข้อมูล",
        "ready": "พร้อม",
        "pending": "รอตรวจสอบ",
        "complete": "เสร็จสิ้น",
    }
    return labels.get(text.casefold(), text)


def _format_display_columns(frame: pd.DataFrame) -> pd.DataFrame:
    display = pd.DataFrame(frame).copy()
    if display.empty:
        return display
    count_terms = ("จำนวน", "count", "vehicles", "คัน", "errors", "warnings", "info", "rows")
    pcu_terms = ("pcu", "PCU")
    percent_terms = ("percent", "share", "สัดส่วน", "%")
    for column in display.columns:
        column_text = str(column)
        lowered = column_text.casefold()
        if any(term.casefold() in lowered for term in pcu_terms):
            display[column] = display[column].map(format_pcu)
        elif "phf" in lowered:
            display[column] = display[column].map(format_phf)
        elif any(term.casefold() in lowered for term in percent_terms):
            display[column] = display[column].map(format_percent)
        elif any(term.casefold() in lowered for term in count_terms):
            display[column] = display[column].map(format_count)
    return display


def _qc_display_frame(qc: pd.DataFrame | None, *, thai_labels: bool = True) -> pd.DataFrame:
    display = pd.DataFrame(qc).copy() if qc is not None else pd.DataFrame()
    if display.empty:
        return display
    preferred = _existing_columns(
        display,
        [
            "file_name",
            "output_stem",
            "severity",
            "category",
            "check",
            "message",
            "detail",
            "affected_field",
            "movement_code",
            "raw_sheet",
        ],
    )
    display = display[preferred] if preferred else display
    if "severity" in display.columns:
        display["severity"] = display["severity"].map(_display_status_label)
    if not thai_labels:
        return display
    return display.rename(
        columns={
            "file_name": "ชื่อไฟล์",
            "output_stem": "ชื่อส่งออก",
            "severity": "ระดับ",
            "category": "หมวด",
            "check": "รายการตรวจ",
            "message": "ข้อความ",
            "detail": "รายละเอียด",
            "affected_field": "ฟิลด์",
            "movement_code": "Movement",
            "raw_sheet": "Sheet ต้นทาง",
        }
    )


def _batch_qc_preview_display_frame(batch_qc: pd.DataFrame) -> pd.DataFrame:
    display = _qc_display_frame(batch_qc, thai_labels=True)
    columns = _existing_columns(
        display,
        ["ชื่อไฟล์", "ชื่อส่งออก", "ระดับ", "หมวด", "รายการตรวจ", "ข้อความ", "รายละเอียด", "Movement", "Sheet ต้นทาง"],
    )
    return display[columns] if columns else display


def _batch_summary_counts(status_frame: pd.DataFrame) -> dict[str, int]:
    if status_frame.empty:
        return {
            "total_files": 0,
            "successful_files": 0,
            "failed_files": 0,
            "QC_errors": 0,
            "QC_warnings": 0,
            "QC_info": 0,
        }
    statuses = status_frame["status"].fillna("").astype(str).str.casefold() if "status" in status_frame else pd.Series(dtype=str)

    def _sum_column(column: str) -> int:
        if column not in status_frame:
            return 0
        return int(pd.to_numeric(status_frame[column], errors="coerce").fillna(0).sum())

    return {
        "total_files": int(len(status_frame)),
        "successful_files": int((statuses == "success").sum()),
        "failed_files": int((statuses == "failed").sum()),
        "QC_errors": _sum_column("QC errors"),
        "QC_warnings": _sum_column("QC warnings"),
        "QC_info": _sum_column("QC info"),
    }


def _existing_columns(frame: pd.DataFrame, columns: list[str]) -> list[str]:
    return [column for column in columns if column in frame.columns]


def _render_peak_card(title: str, period_label: str, pcu: str, source: str) -> None:
    is_confirmed = source == "user_confirmed"
    card_class = "tmc-peak-confirmed" if is_confirmed else "tmc-peak-suggested"
    title_text = str(title).casefold()
    if "am" in title_text or "เช้า" in title_text:
        card_class = f"{card_class} tmc-peak-am"
    elif "pm" in title_text or "เย็น" in title_text:
        card_class = f"{card_class} tmc-peak-pm"
    badge = "กำหนดแล้ว" if is_confirmed else "แนะนำ"
    pcu_text = f"{pcu} PCU" if pcu else "ไม่มีข้อมูล PCU"
    st.markdown(
        f'<div class="tmc-card tmc-peak-card {card_class}">'
        '<div class="tmc-peak-topline">'
        f'<div class="tmc-card-label">{escape(title)}</div>'
        f'<div class="tmc-peak-badge">{badge}</div>'
        "</div>"
        f'<div class="tmc-peak-time">{escape(period_label or "ไม่มีข้อมูล")}</div>'
        f'<div class="tmc-peak-pcu">{escape(pcu_text)}</div>'
        "</div>",
        unsafe_allow_html=True,
    )


def _readiness_item(label: str, ready: bool, detail: str = "") -> None:
    st.markdown(_readiness_item_html(label, ready, detail), unsafe_allow_html=True)


def _readiness_item_html(label: str, ready: bool, detail: str = "") -> str:
    status = "พร้อม" if ready else "ต้องตรวจสอบ"
    css_class = "tmc-check-ready" if ready else "tmc-check-warn"
    icon = "✓" if ready else "!"
    detail_text = f'<div class="tmc-check-detail">{escape(detail)}</div>' if detail else ""
    return (
        f'<div class="tmc-check-item {css_class}">'
        f'<span class="tmc-check-icon">{icon}</span>'
        '<span class="tmc-check-body">'
        f'<div class="tmc-check-label">{escape(label)}</div>'
        f'<div class="tmc-check-status">{status}</div>'
        f"{detail_text}"
        "</span>"
        "</div>"
    )


def _render_readiness_checklist(items: list[tuple[str, bool, str]]) -> None:
    st.markdown(
        '<div class="tmc-checklist">'
        + "".join(_readiness_item_html(label, ready, detail) for label, ready, detail in items)
        + "</div>",
        unsafe_allow_html=True,
    )


def _qc_severity_counts(qc: pd.DataFrame | None) -> dict[str, int]:
    counts = {"error": 0, "warning": 0, "info": 0}
    if qc is None or qc.empty or "severity" not in qc.columns:
        return counts
    severity_counts = qc["severity"].fillna("").astype(str).str.casefold().value_counts()
    for severity in counts:
        counts[severity] = int(severity_counts.get(severity, 0))
    return counts


def _qc_status_text(qc: pd.DataFrame | None) -> tuple[str, str, str]:
    counts = _qc_severity_counts(qc)
    if counts["error"]:
        return "Errors found", f"{counts['error']} errors / {counts['warning']} warnings / {counts['info']} info", "red"
    if counts["warning"] or counts["info"]:
        return f"{counts['warning']} warnings / {counts['info']} info", "Non-blocking QC notes", "amber"
    return "No QC issues", "Ready for export review", "green"


def _render_qc_status(qc: pd.DataFrame | None, label: str = "QC status") -> None:
    status, note, tone = _qc_status_text(qc)
    st.markdown(_status_card_html(label, status, note, tone=tone), unsafe_allow_html=True)


def _mapping_issue_display(issues: pd.DataFrame) -> pd.DataFrame:
    if issues.empty:
        return issues
    display = issues.copy()
    if "message" in display.columns and "field" in display.columns:
        display["message"] = display["field"].map(lambda field: f"ต้องระบุ {field} ก่อนประมวลผล")
    return display.rename(
        columns={
            "raw_sheet": "Sheet ต้นทาง",
            "field": "ฟิลด์",
            "message": "รายละเอียด",
        }
    )


def _thai_aggregation_message(message: str) -> str:
    marker = " is aggregated from "
    suffix = " source streams."
    if marker in message and message.endswith(suffix):
        movement_code, count_text = message.removesuffix(suffix).split(marker, 1)
        return f"{movement_code} รวมจาก source stream จำนวน {count_text} รายการ"
    return message


def _mapping_workspace_counts(mapping: pd.DataFrame, detected_sheet_names: list[str] | tuple[str, ...]) -> dict[str, int]:
    frame = pd.DataFrame(mapping)
    if frame.empty:
        return {
            "rows": 0,
            "detected_sheets": len(detected_sheet_names or []),
            "included": 0,
            "excluded": 0,
            "peak_included": 0,
            "duplicate_movements": 0,
            "blank_source_stream": 0,
        }
    include_report = frame["include_in_report"].fillna(True).astype(bool) if "include_in_report" in frame else pd.Series(True, index=frame.index)
    include_peak = frame["include_in_peak"].fillna(True).astype(bool) if "include_in_peak" in frame else pd.Series(True, index=frame.index)
    movement = frame["movement_code"].fillna("").astype(str).str.strip() if "movement_code" in frame else pd.Series("", index=frame.index)
    included_movement = movement[include_report & (movement != "")]
    source_stream = frame["source_stream"].fillna("").astype(str).str.strip() if "source_stream" in frame else pd.Series("", index=frame.index)
    return {
        "rows": int(len(frame)),
        "detected_sheets": len(detected_sheet_names or []),
        "included": int(include_report.sum()),
        "excluded": int((~include_report).sum()),
        "peak_included": int(include_peak.sum()),
        "duplicate_movements": int((included_movement.value_counts() > 1).sum()),
        "blank_source_stream": int((source_stream == "").sum()),
    }


def _mapping_aggregation_preview(mapping: pd.DataFrame) -> pd.DataFrame:
    frame = pd.DataFrame(mapping)
    columns = ["output_movement_code", "source_stream count", "raw_movement_label/source_direction", "aggregation_method"]
    if frame.empty or "movement_code" not in frame:
        return pd.DataFrame(columns=columns)
    include_report = frame["include_in_report"].fillna(True).astype(bool) if "include_in_report" in frame else pd.Series(True, index=frame.index)
    movement = frame["movement_code"].fillna("").astype(str).str.strip()
    included = frame[include_report & (movement != "")].copy()
    if included.empty:
        return pd.DataFrame(columns=columns)
    included["output_movement_code"] = included["movement_code"].fillna("").astype(str).str.strip()
    for column in ("source_stream", "raw_movement_label", "raw_direction", "raw_sheet", "aggregation_method"):
        if column not in included:
            included[column] = ""
        included[column] = included[column].fillna("").astype(str)
    rows = []
    for movement_code, group in included.groupby("output_movement_code", dropna=False):
        if len(group) <= 1:
            continue
        label_parts = []
        for _, row in group.iterrows():
            label = str(row.get("raw_movement_label") or row.get("raw_direction") or row.get("raw_sheet") or "").strip()
            if label:
                label_parts.append(label)
        rows.append(
            {
                "output_movement_code": movement_code,
                "source_stream count": int(group["source_stream"].replace("", pd.NA).nunique(dropna=True) or len(group)),
                "raw_movement_label/source_direction": ", ".join(label_parts[:4]),
                "aggregation_method": ", ".join(sorted({str(value) for value in group["aggregation_method"] if str(value).strip()})),
            }
        )
    return pd.DataFrame(rows, columns=columns)


def _mapping_preset_rows_frame(preset: dict[str, object] | None) -> pd.DataFrame:
    if not preset:
        return pd.DataFrame()
    return pd.DataFrame(preset.get("mapping_rows") or preset.get("rows") or [])


def _current_mapping_scheme() -> str:
    return _movement_scheme(str(st.session_state.get("tmc_mapping_code_scheme") or MOVEMENT_SCHEME_V1))


def _set_current_mapping_scheme(scheme: str) -> None:
    st.session_state["tmc_mapping_code_scheme"] = _movement_scheme(str(scheme or MOVEMENT_SCHEME_V1))


def _render_mapping_scheme_status(scheme: str) -> None:
    _render_alert(f"movement_code_scheme: {scheme} / {movement_scheme_description(scheme)}", "info")
    if _is_v2_scheme(scheme):
        _render_alert(f"movement_code_scheme: {scheme} / supported for single-file workflow", "success")
        _render_alert("Batch approach_movement รองรับ Safe PNG/generated mode; Excel Template Mode ยังถูกบล็อก", "info")
    elif mapping_is_process_compatible(scheme):
        _render_alert(f"movement_code_scheme: {scheme} / compatible", "success")
    else:
        _render_alert(f"movement_code_scheme: {scheme} / loaded and valid, not process-compatible yet", "warning")
        _render_alert(mapping_processing_block_reason(scheme), "warning")


def _thai_mapping_control_warning(message: str) -> str:
    if message.startswith("Unknown ") and " value(s) " in message:
        return f"พบค่าใน Mapping ที่อยู่นอกตัวเลือก ระบบโหลดแบบปลอดภัยแล้ว ({message})"
    if message.startswith("Legacy movement code value(s)"):
        return f"พบ movement_code จากไฟล์เดิม ระบบคงค่าไว้เพื่อความเข้ากันได้ ({message})"
    return message


def _render_template_audit_notes() -> None:
    try:
        audit = audit_template_formulas(DEFAULT_TEMPLATE_PATH, load_template_map(), "Summary")
    except Exception as exc:  # pragma: no cover - UI guardrail
        st.warning(f"ไม่สามารถโหลดรายละเอียดเทมเพลตได้: {exc}")
        return
    summary = {
        "formula_cells": len(audit.formula_cells),
        "mapped_formula_cells": len(audit.mapped_formula_cells),
        "external_links": len(audit.external_links),
        "ref_errors": len(audit.ref_errors),
        "missing_sheet_issues": len(audit.missing_sheet_issues),
    }
    st.write(summary)
    warning_rows = [
        {"type": "external_link", "sheet": item.sheet, "cell": item.cell, "formula": item.formula}
        for item in audit.external_links
    ]
    warning_rows.extend(
        {"type": "ref_error", "sheet": item.sheet, "cell": item.cell, "formula": item.formula}
        for item in audit.ref_errors
    )
    warning_rows.extend(
        {
            "type": "missing_sheet",
            "sheet": item.sheet,
            "cell": item.cell,
            "formula": item.formula,
            "referenced_sheet": item.referenced_sheet,
        }
        for item in audit.missing_sheet_issues
    )
    if warning_rows:
        st.dataframe(pd.DataFrame(warning_rows), width="stretch")
    else:
        st.caption("ไม่พบคำเตือนสูตรเทมเพลตที่เป็นประเด็นสำคัญ")

def _run_streamlit_app() -> None:
    st.set_page_config(page_title="TMC Processor", layout="wide")
    _inject_global_css()

    for _key in list(st.session_state.keys()):
        if _key.startswith("excel_com_probe") or _key.startswith("tmc_excel_com_probe"):
            st.session_state.pop(_key, None)

    excel_com_status = _probe_excel_com_for_ui()

    st.session_state.setdefault("work_mode", "ประมวลผลไฟล์เดียว")
    work_mode_options = ["ประมวลผลไฟล์เดียว", "ประมวลผลหลายไฟล์"]
    if st.session_state.get("work_mode") not in work_mode_options:
        st.session_state["work_mode"] = "ประมวลผลไฟล์เดียว"
    is_single_file_mode = st.session_state["work_mode"] == "ประมวลผลไฟล์เดียว"
    uploaded_file = None
    batch_uploads = []
    batch_preset_upload = None

    with st.sidebar:
        _render_sidebar_brand()
        _render_sidebar_section("Work mode")
        work_mode = st.radio(
            "เลือกโหมดการทำงาน",
            options=work_mode_options,
            key="work_mode",
            label_visibility="collapsed",
        )
        is_single_file_mode = work_mode == "ประมวลผลไฟล์เดียว"
        _render_sidebar_section("Input")
        if is_single_file_mode:
            st.caption("ใช้สำหรับประมวลผลไฟล์ TMC หนึ่งไฟล์ และตรวจ Peak ก่อนส่งออกรายงาน")
            uploaded_file = st.file_uploader(
                "อัปโหลดไฟล์ TMC Excel",
                type=["xlsx", "xlsm", "xls"],
                key="raw_tmc_upload",
            )
        else:
            st.caption("เหมาะสำหรับจุดสำรวจเดียวกันหลายวัน โดยใช้ Mapping Preset เดียวกันทุกไฟล์")
            batch_uploads = st.file_uploader(
                "อัปโหลดไฟล์ TMC Excel หลายไฟล์",
                type=["xlsx", "xlsm", "xls"],
                accept_multiple_files=True,
                key="batch_raw_tmc_uploads",
                help="อัปโหลด raw TMC workbook หลายไฟล์ของจุดสำรวจเดียวกัน เช่น หลายวันสำรวจ",
            )
            batch_preset_upload = st.file_uploader(
                "เปิด Mapping Preset สำหรับทุกไฟล์",
                type=["json"],
                key="batch_mapping_preset_upload",
                help="Batch ใช้ Mapping Preset เดียวกันกับทุกไฟล์",
            )

    file_bytes = uploaded_file.getvalue() if uploaded_file is not None else b""
    uploaded_identity = (uploaded_file.name, len(file_bytes)) if uploaded_file is not None else None
    uploaded_content_revision = source_fingerprint(file_bytes) if uploaded_file is not None else None
    uploaded_filename_default = _default_text_from_filename(uploaded_file.name if uploaded_file is not None else None)
    initialize_setup_state_once(uploaded_filename_default)

    if uploaded_identity and (
        st.session_state.get("tmc_uploaded_identity") != uploaded_identity
        or st.session_state.get("tmc_uploaded_source_fingerprint") != uploaded_content_revision
    ):
        st.session_state.pop("tmc_output", None)
        st.session_state.pop("tmc_processed", None)
        st.session_state.pop("tmc_pce_results_stale", None)
        if not st.session_state.get("tmc_mapping_table_from_session"):
            st.session_state.pop("mapping_table", None)
            _set_current_mapping_scheme(MOVEMENT_SCHEME_V1)
            _clear_mapping_for_scheme_change()
        for key in list(st.session_state.keys()):
            if str(key).startswith("mapping_editor_"):
                st.session_state.pop(key, None)
        for key in ["am_peak_period_select", "pm_peak_period_select"]:
            st.session_state.pop(key, None)
        st.session_state["tmc_uploaded_identity"] = uploaded_identity
        st.session_state["tmc_uploaded_source_fingerprint"] = uploaded_content_revision
        apply_uploaded_filename_defaults_to_setup_state(uploaded_filename_default)
    _hydrate_setup_widgets_from_state()

    st.session_state.setdefault("mapping_editor_version", 0)
    st.session_state.setdefault("mapping_editor_view_mode", "Basic")
    st.session_state.setdefault("tmc_mapping_code_scheme", MOVEMENT_SCHEME_V1)
    st.session_state.setdefault("tmc_mapping_source", MAPPING_SOURCE_DEFAULT_PREVIEW)
    st.session_state.setdefault("tmc_batch_file_metadata_table", [])
    st.session_state.setdefault("tmc_batch_file_metadata_editor_version", 0)
    st.session_state.setdefault("tmc_batch_analysis_result", None)
    st.session_state.setdefault("tmc_batch_selected_review_file", "")
    st.session_state.setdefault("tmc_batch_confirmed_peaks", {})
    st.session_state.setdefault("tmc_batch_export_result", None)
    st.session_state.setdefault("tmc_batch_export_mode", None)
    _ensure_pce_factor_state()
    single_export_options = _single_export_mode_options(excel_com_status)
    export_mode = _coerce_export_mode(
        st.session_state.get("report_export_mode"),
        single_export_options,
        _default_single_export_mode(excel_com_status),
    )
    st.session_state["report_export_mode"] = export_mode
    use_excel_com_native_charts = _use_excel_native_charts_for_export(export_mode, excel_com_status)
    batch_export_options = _batch_export_mode_options(excel_com_status)
    st.session_state["tmc_batch_export_mode"] = _coerce_export_mode(
        st.session_state.get("tmc_batch_export_mode"),
        batch_export_options,
        _default_batch_export_mode(excel_com_status),
    )

    if is_single_file_mode:
        _sync_single_workflow_from_state(
            source_bytes=file_bytes if uploaded_file is not None else None,
            source_file_name=uploaded_file.name if uploaded_file is not None else None,
            export_mode=export_mode,
        )

    with st.sidebar:
        if is_single_file_mode:
            _render_sidebar_section("Project session")
            _render_project_session_section(
                uploaded_file.name if uploaded_file is not None else None,
                len(file_bytes) if uploaded_file is not None else None,
                compact=True,
            )
        else:
            use_excel_com_native_charts = _use_excel_native_charts_for_export(export_mode, excel_com_status)
        _render_sidebar_section("Engine status")
        previous_excel_com_available = bool(excel_com_status.available)
        if st.button("ทดสอบ Excel COM", key="test_excel_com"):
            excel_com_status = _probe_excel_com_for_ui(force=True)
            if bool(excel_com_status.available) != previous_excel_com_available:
                st.rerun()

        if excel_com_status.available:
            version_text = f"Excel version: {excel_com_status.version}" if excel_com_status.version else "พร้อมใช้งาน"
            _render_sidebar_badge("Excel COM พร้อมใช้งาน", version_text, ready=True)
        else:
            detail = f"{excel_com_status.reason}"
            if excel_com_status.detail:
                detail = f"{detail} · {excel_com_status.detail}"
            _render_sidebar_badge("Excel COM ไม่พร้อมใช้งาน", "ระบบจะใช้โหมดสำรองแบบ PNG", ready=False)
            st.caption(detail)
        with st.expander("รายละเอียด Excel COM", expanded=False):
            _render_excel_com_status(excel_com_status)
        _render_version_stamp()

    batch_metadata_rows = _sync_batch_metadata_state(batch_uploads) if not is_single_file_mode else []
    loaded_batch_preset = None
    batch_preset_name = ""
    batch_preset_signature = None
    batch_mapping_scheme = MOVEMENT_SCHEME_V1
    if not is_single_file_mode and batch_preset_upload is not None:
        try:
            batch_preset_bytes = batch_preset_upload.getvalue()
            batch_preset_signature = (batch_preset_upload.name, hashlib.sha256(batch_preset_bytes).hexdigest())
            loaded = load_mapping_preset(batch_preset_bytes)
            loaded_batch_preset = loaded.preset
            batch_mapping_scheme = detect_mapping_preset_scheme(loaded)
            batch_preset_signature = semantic_fingerprint(
                {
                    "movement_code_scheme": batch_mapping_scheme,
                    "mapping_rows": loaded_batch_preset.get("mapping_rows", []),
                }
            )
            st.session_state["tmc_batch_mapping_code_scheme"] = batch_mapping_scheme
            batch_export_options = _batch_export_mode_options(excel_com_status, batch_mapping_scheme)
            st.session_state["tmc_batch_export_mode"] = _coerce_export_mode(
                st.session_state.get("tmc_batch_export_mode"),
                batch_export_options,
                _default_batch_export_mode(excel_com_status, batch_mapping_scheme),
            )
            batch_preset_name = str(loaded.preset.get("preset_name") or batch_preset_upload.name)
            for warning_message in loaded.warnings:
                st.sidebar.warning(warning_message)
        except (MappingPresetError, ValueError) as exc:
            st.sidebar.error(f"ไม่สามารถเปิด Mapping Preset ได้: {exc}")

    if not is_single_file_mode:
        _sync_batch_workflow_from_state(
            batch_uploads=batch_uploads,
            mapping_preset=loaded_batch_preset,
            movement_code_scheme=batch_mapping_scheme,
            metadata_rows=batch_metadata_rows,
            export_mode=str(st.session_state.get("tmc_batch_export_mode") or BATCH_SAFE_PNG_EXPORT_LABEL),
        )

    _render_app_header()
    top_status_placeholder = st.empty()
    workflow_shell_placeholder = st.empty()
    shell_batch_export_mode = st.session_state.get("tmc_batch_export_mode", BATCH_SAFE_PNG_EXPORT_LABEL)
    shell_batch_signature = _batch_analysis_signature(
        uploads_signature=_batch_upload_signature(batch_uploads),
        preset_signature=batch_preset_signature,
        pce_factors=_current_pce_factors_from_state(),
        peak_mode=st.session_state.get("peak_mode_select", DEFAULT_PEAK_MODE),
        peak_windows={
            "AM": (
                _time_text(st.session_state.get("am_peak_window_start_input", _time_from_text(AM_WINDOW[0]))),
                _time_text(st.session_state.get("am_peak_window_end_input", _time_from_text(AM_WINDOW[1]))),
            ),
            "PM": (
                _time_text(st.session_state.get("pm_peak_window_start_input", _time_from_text(PM_WINDOW[0]))),
                _time_text(st.session_state.get("pm_peak_window_end_input", _time_from_text(PM_WINDOW[1]))),
            ),
        },
    )
    flash_message = st.session_state.pop("tmc_flash_message", None)
    if flash_message:
        _render_alert(str(flash_message.get("message", "")), str(flash_message.get("kind", "success")))

    detected_sheet_names: list[str] = []
    preview_summary = pd.DataFrame()
    previews: dict[str, pd.DataFrame] = {}
    parsed_details = {}
    if uploaded_file is not None:
        excel_buffer = BytesIO(file_bytes)
        try:
            detected_sheet_names = detect_raw_direction_sheet_names(excel_buffer)
            st.session_state["tmc_detected_sheet_names"] = list(detected_sheet_names)
            excel_buffer.seek(0)
            preview_summary, previews = preview_detected_sheets(excel_buffer)
            excel_buffer.seek(0)
            parsed_details = load_detected_sheet_details(excel_buffer)
        except Exception as exc:  # pragma: no cover - UI guardrail
            st.error(f"ไม่สามารถอ่านไฟล์ Workbook ได้: {exc}")
            st.stop()

    active_tab = render_workflow_navigation()

    project_name = st.session_state.get("project_name_input", "")
    tmc_id = st.session_state.get("tmc_id_input", "")
    tmc_title = st.session_state.get("tmc_title_input", uploaded_filename_default)
    survey_point = st.session_state.get("survey_point_input", uploaded_filename_default)
    survey_date_text = st.session_state.get("survey_date_text_input", "")
    weather = st.session_state.get("weather_input", DEFAULT_WEATHER)
    responsible_party = st.session_state.get("responsible_party_input", DEFAULT_RESPONSIBLE_PARTY)
    survey_period = st.session_state.get("survey_period_input", DEFAULT_SURVEY_PERIOD)
    north_label = st.session_state.get("north_label_input", "")
    south_label = st.session_state.get("south_label_input", "")
    east_label = st.session_state.get("east_label_input", "")
    west_label = st.session_state.get("west_label_input", "")
    north_road = st.session_state.get("north_road_input", "")
    south_road = st.session_state.get("south_road_input", "")
    east_road = st.session_state.get("east_road_input", "")
    west_road = st.session_state.get("west_road_input", "")
    caption_text = st.session_state.get("caption_text_input", DEFAULT_CAPTION_TEXT)
    peak_mode = st.session_state.get("peak_mode_select", DEFAULT_PEAK_MODE)
    am_peak_window_start = st.session_state.get("am_peak_window_start_input", _time_from_text(AM_WINDOW[0]))
    am_peak_window_end = st.session_state.get("am_peak_window_end_input", _time_from_text(AM_WINDOW[1]))
    pm_peak_window_start = st.session_state.get("pm_peak_window_start_input", _time_from_text(PM_WINDOW[0]))
    pm_peak_window_end = st.session_state.get("pm_peak_window_end_input", _time_from_text(PM_WINDOW[1]))
    show_u_turn = st.session_state.get("show_u_turn_checkbox", True)
    selected_pce_factors = _current_pce_factors_from_state()

    if is_single_file_mode:
      if active_tab == "ตั้งค่า":
        _render_section_header("ตั้งค่างาน", "ระบุข้อมูลงานและค่าที่ใช้ในการประมวลผลรายงาน")
        if uploaded_file is None:
            _render_action_hint("เริ่มจากอัปโหลดไฟล์ TMC Excel ที่แถบด้านซ้าย")

        setup_left, setup_right = st.columns([1.15, 1])
        with setup_left:
            with st.container(border=True):
                _render_section_header("ข้อมูลโครงการและรายงาน", "ข้อมูลหลักสำหรับปกและหัวรายงาน")
                report_cols = st.columns(2)
                project_name = report_cols[0].text_input("ชื่อโครงการ", key="project_name_input")
                tmc_id = report_cols[1].text_input("TMC ID", key="tmc_id_input")
                tmc_title = st.text_input("ชื่อจุดนับ", key="tmc_title_input")
                info_cols = st.columns(2)
                survey_point = info_cols[0].text_input("จุดสำรวจ", key="survey_point_input")
                survey_date_text = info_cols[1].text_input("วันที่สำรวจ", key="survey_date_text_input")
                weather = info_cols[0].text_input("สภาพอากาศ", key="weather_input")
                responsible_party = info_cols[1].text_input("ผู้รับผิดชอบ", key="responsible_party_input")

            with st.container(border=True):
                _render_section_header("ช่วงสำรวจและ Peak", "กำหนดกรอบเวลาที่ใช้คัดเลือก Peak อัตโนมัติ")
                survey_period = st.text_input("ช่วงเวลาสำรวจ", key="survey_period_input")
                if st.session_state.get("peak_mode_select", DEFAULT_PEAK_MODE) not in PEAK_MODE_OPTIONS:
                    st.session_state["peak_mode_select"] = DEFAULT_PEAK_MODE
                peak_mode = st.selectbox("รูปแบบการคำนวณ Peak", options=PEAK_MODE_OPTIONS, key="peak_mode_select")
                period_cols = st.columns(4)
                am_peak_window_start = period_cols[0].time_input("เริ่มช่วง AM", step=900, key="am_peak_window_start_input")
                am_peak_window_end = period_cols[1].time_input("สิ้นสุดช่วง AM", step=900, key="am_peak_window_end_input")
                pm_peak_window_start = period_cols[2].time_input("เริ่มช่วง PM", step=900, key="pm_peak_window_start_input")
                pm_peak_window_end = period_cols[3].time_input("สิ้นสุดช่วง PM", step=900, key="pm_peak_window_end_input")

        with setup_right:
            with st.container(border=True):
                _render_section_header("ป้ายปลายทางและถนน", "ข้อความที่ใช้ใน Diagram และรายงาน")
                st.caption("ป้ายปลายทาง")
                direction_cols = st.columns(2)
                north_label = direction_cols[0].text_input("ป้ายปลายทางด้านเหนือ", key="north_label_input")
                south_label = direction_cols[1].text_input("ป้ายปลายทางด้านใต้", key="south_label_input")
                east_label = direction_cols[0].text_input("ป้ายปลายทางด้านตะวันออก", key="east_label_input")
                west_label = direction_cols[1].text_input("ป้ายปลายทางด้านตะวันตก", key="west_label_input")
                st.caption("ชื่อถนน / ทางหลวง")
                road_cols = st.columns(2)
                north_road = road_cols[0].text_input("ชื่อถนนด้านเหนือ", key="north_road_input")
                south_road = road_cols[1].text_input("ชื่อถนนด้านใต้", key="south_road_input")
                east_road = road_cols[0].text_input("ชื่อถนนด้านตะวันออก", key="east_road_input")
                west_road = road_cols[1].text_input("ชื่อถนนด้านตะวันตก", key="west_road_input")
                caption_text = st.text_input("คำบรรยายรูป Diagram", key="caption_text_input")
                show_u_turn = st.checkbox("แสดง movement กลับรถ", key="show_u_turn_checkbox")

            with st.container(border=True):
                _render_section_header("ค่า PCE", "แก้เฉพาะกรณีต้องใช้ค่าเทียบเท่ารถยนต์นั่งต่างจากค่าเริ่มต้น")
                selected_pce_factors = _render_pce_factor_editor()
                _sync_workflow_after_pce_editor(
                    is_single_file_mode=True,
                    source_bytes=file_bytes if uploaded_file is not None else None,
                    source_file_name=uploaded_file.name if uploaded_file is not None else None,
                    export_mode=export_mode,
                )
                has_overrides, override_text = _pce_override_summary(selected_pce_factors)
                _render_status_chip("มีค่า PCE ที่แก้ไขเอง" if has_overrides else "ใช้ค่า PCE เริ่มต้น", "warning" if has_overrides else "success")
                if has_overrides:
                    st.caption(override_text)

    if not is_single_file_mode:
        project_name = st.session_state.get("project_name_input", "")
        tmc_id = st.session_state.get("tmc_id_input", "")
        tmc_title = st.session_state.get("tmc_title_input", "")
        survey_point = st.session_state.get("survey_point_input", "")
        survey_date_text = st.session_state.get("survey_date_text_input", "")
        weather = st.session_state.get("weather_input", DEFAULT_WEATHER)
        responsible_party = st.session_state.get("responsible_party_input", DEFAULT_RESPONSIBLE_PARTY)
        survey_period = st.session_state.get("survey_period_input", DEFAULT_SURVEY_PERIOD)
        north_label = st.session_state.get("north_label_input", "")
        south_label = st.session_state.get("south_label_input", "")
        east_label = st.session_state.get("east_label_input", "")
        west_label = st.session_state.get("west_label_input", "")
        north_road = st.session_state.get("north_road_input", "")
        south_road = st.session_state.get("south_road_input", "")
        east_road = st.session_state.get("east_road_input", "")
        west_road = st.session_state.get("west_road_input", "")
        caption_text = st.session_state.get("caption_text_input", DEFAULT_CAPTION_TEXT)
        peak_mode = st.session_state.get("peak_mode_select", DEFAULT_PEAK_MODE)
        am_peak_window_start = st.session_state.get("am_peak_window_start_input", _time_from_text(AM_WINDOW[0]))
        am_peak_window_end = st.session_state.get("am_peak_window_end_input", _time_from_text(AM_WINDOW[1]))
        pm_peak_window_start = st.session_state.get("pm_peak_window_start_input", _time_from_text(PM_WINDOW[0]))
        pm_peak_window_end = st.session_state.get("pm_peak_window_end_input", _time_from_text(PM_WINDOW[1]))
        show_u_turn = st.session_state.get("show_u_turn_checkbox", True)
        selected_pce_factors = _current_pce_factors_from_state()

    setup = build_setup_for_processing(uploaded_file.name if uploaded_file is not None else "")
    setup["movement_code_scheme"] = _current_mapping_scheme()
    project_name = str(setup.get("project_name", "") or "")
    tmc_id = str(setup.get("tmc_id", "") or "")
    tmc_title = str(setup.get("tmc_title", "") or "")
    survey_point = str(setup.get("survey_point", "") or "")
    survey_date_text = str(setup.get("survey_date_text", "") or "")
    weather = str(setup.get("weather", DEFAULT_WEATHER) or "")
    responsible_party = str(setup.get("responsible_party", DEFAULT_RESPONSIBLE_PARTY) or "")
    survey_period = str(setup.get("survey_period", DEFAULT_SURVEY_PERIOD) or "")
    north_label = str(setup.get("north_label", "") or "")
    south_label = str(setup.get("south_label", "") or "")
    east_label = str(setup.get("east_label", "") or "")
    west_label = str(setup.get("west_label", "") or "")
    north_road = str(setup.get("north_road", "") or "")
    south_road = str(setup.get("south_road", "") or "")
    east_road = str(setup.get("east_road", "") or "")
    west_road = str(setup.get("west_road", "") or "")
    caption_text = str(setup.get("caption_text", DEFAULT_CAPTION_TEXT) or "")
    peak_mode = str(setup.get("peak_mode", DEFAULT_PEAK_MODE) or DEFAULT_PEAK_MODE)
    am_peak_window_start = _coerce_setup_time(setup.get("am_peak_window_start"), _time_from_text(AM_WINDOW[0]))
    am_peak_window_end = _coerce_setup_time(setup.get("am_peak_window_end"), _time_from_text(AM_WINDOW[1]))
    pm_peak_window_start = _coerce_setup_time(setup.get("pm_peak_window_start"), _time_from_text(PM_WINDOW[0]))
    pm_peak_window_end = _coerce_setup_time(setup.get("pm_peak_window_end"), _time_from_text(PM_WINDOW[1]))
    show_u_turn = bool(setup.get("show_u_turn", True))
    peak_windows = {
        "AM": (setup["am_peak_window_start"], setup["am_peak_window_end"]),
        "PM": (setup["pm_peak_window_start"], setup["pm_peak_window_end"]),
    }
    mapping = pd.DataFrame(st.session_state.get("mapping_table") or [])

    if is_single_file_mode:
        if active_tab == "กำหนดทิศทาง":
            _render_section_header(
                "กำหนดทิศทาง",
                "จับคู่ข้อมูลจากชีตสำรวจเข้ากับรหัสการเคลื่อนที่มาตรฐาน ก่อนประมวลผลรายงาน",
            )
            if uploaded_file is None:
                _render_empty_state(
                    "ยังไม่มีไฟล์สำรวจ",
                    "เริ่มจากอัปโหลดไฟล์ TMC Excel ที่แถบด้านซ้าย",
                )
            elif not detected_sheet_names:
                _render_alert('ไม่พบ Sheet ทิศทางจากไฟล์สำรวจ ควรมีชื่อ Sheet เช่น "ทิศ 1", "ทิศ 2", หรือ "ทิศ 2+3"', "warning")
            else:
                default_mapping = default_mapping_for_sheets(detected_sheet_names)
                if st.session_state.get("mapping_table") is not None:
                    default_mapping = apply_saved_mapping_to_sheets(detected_sheet_names, pd.DataFrame(st.session_state["mapping_table"]))
                mapping_scheme = _current_mapping_scheme()
                mapping_rows_committed = _mapping_rows_are_committed()
                selected_scheme = st.selectbox(
                    "ระบบรหัส Movement",
                    options=MOVEMENT_SCHEMES,
                    index=MOVEMENT_SCHEMES.index(mapping_scheme),
                    format_func=_scheme_select_label,
                    disabled=mapping_rows_committed,
                    key="movement_code_scheme_selector",
                    help="เลือกได้ก่อนโหลดหรือกรอก Mapping; ถ้ามี Mapping แล้วให้ล้างหรือโหลด Mapping ใหม่เพื่อไม่ตีความรหัสเดิมผิด",
                )
                if selected_scheme != mapping_scheme and not mapping_rows_committed:
                    _set_current_mapping_scheme(selected_scheme)
                    mapping_scheme = selected_scheme
                    st.session_state["mapping_editor_version"] = int(st.session_state.get("mapping_editor_version", 0) or 0) + 1
                if mapping_rows_committed:
                    st.caption(f"Detected movement_code_scheme: {mapping_scheme}")
                    st.caption(MAPPING_SCHEME_LOCK_CAPTION)
                    if st.button("ล้าง Mapping เพื่อเปลี่ยนระบบรหัส", key="clear_mapping_for_scheme_change"):
                        _clear_mapping_for_scheme_change()
                        st.rerun()
                else:
                    st.caption(f"Selected movement_code_scheme: {mapping_scheme}")
                preset_name_seed = st.session_state.get("tmc_id_input") or st.session_state.get("tmc_title_input") or uploaded_file.name
                preset_source = pd.DataFrame(st.session_state.get("mapping_table") or default_mapping.to_dict("records"))
                preset_bytes = serialize_mapping_preset(
                    build_mapping_preset(
                        preset_source,
                        preset_name=str(preset_name_seed or "TMC Mapping Preset"),
                        movement_code_scheme=mapping_scheme,
                    )
                )
                st.session_state["tmc_mapping_preset_bytes"] = preset_bytes
                st.session_state["tmc_mapping_preset_filename"] = safe_mapping_preset_filename(preset_name_seed)

                with st.container(border=True):
                    _render_section_header("นำเข้า/ส่งออก Mapping", "เลือกใช้ Mapping Excel สำหรับแก้ไขใน Excel หรือ Mapping Preset สำหรับนำค่าที่ตั้งไว้กลับมาใช้ซ้ำ")
                    mapping_excel_col, mapping_preset_col = st.columns(2)
                    with mapping_excel_col:
                        st.markdown("**Mapping Excel**")
                        st.caption("สำหรับกรอกหรือแก้ไข Mapping ด้วย Excel")
                        mapping_upload = st.file_uploader(
                            "โหลดไฟล์ Mapping Excel",
                            type=["xlsx", "xlsm", "xls"],
                            key="mapping_upload",
                        )
                        _render_download_button(
                            "ดาวน์โหลดเทมเพลต Mapping",
                            mapping_to_excel_bytes(default_mapping, movement_code_scheme=mapping_scheme),
                            "tmc_mapping_template.xlsx",
                            EXCEL_MIME,
                        )
                    with mapping_preset_col:
                        st.markdown("**Mapping Preset**")
                        st.caption("สำหรับบันทึก Mapping ที่ตั้งค่าแล้วและนำกลับมาใช้ซ้ำในโปรแกรม")
                        mapping_preset_upload = st.file_uploader(
                            "เปิด Mapping Preset",
                            type=["json"],
                            key="mapping_preset_upload",
                        )
                        st.download_button(
                            "ดาวน์โหลด Mapping Preset",
                            data=download_buffer(preset_bytes),
                            file_name=st.session_state["tmc_mapping_preset_filename"],
                            mime=MAPPING_PRESET_MIME,
                            key="download_mapping_preset",
                        )
                if mapping_upload is not None:
                    try:
                        mapping_upload_bytes = mapping_upload.getvalue()
                        mapping_upload_identity = (mapping_upload.name, hashlib.sha256(mapping_upload_bytes).hexdigest())
                        if (
                            st.session_state.get("tmc_mapping_upload_identity") != mapping_upload_identity
                            and st.session_state.get("tmc_mapping_upload_ignored_identity") != mapping_upload_identity
                        ):
                            st.session_state.pop("tmc_mapping_upload_ignored_identity", None)
                            loaded_excel = read_mapping_excel_with_metadata(BytesIO(mapping_upload_bytes))
                            default_mapping = apply_saved_mapping_to_sheets(detected_sheet_names, loaded_excel.mapping)
                            _set_current_mapping_scheme(loaded_excel.movement_code_scheme)
                            st.session_state["tmc_loaded_mapping_code_scheme"] = loaded_excel.movement_code_scheme
                            st.session_state["mapping_table"] = default_mapping.to_dict("records")
                            st.session_state["mapping_editor_version"] = int(st.session_state.get("mapping_editor_version", 0) or 0) + 1
                            st.session_state["tmc_mapping_table_from_session"] = False
                            _set_mapping_source(MAPPING_SOURCE_MAPPING_EXCEL)
                            st.session_state["tmc_mapping_upload_identity"] = mapping_upload_identity
                            _render_alert("โหลด Mapping และปรับใช้กับ Sheet ที่ตรวจพบแล้ว", "success")
                    except Exception as exc:  # pragma: no cover - UI guardrail
                        st.error(f"ไม่สามารถโหลดไฟล์ Mapping ได้: {exc}")
    
                if mapping_preset_upload is not None:
                    try:
                        preset_upload_bytes = mapping_preset_upload.getvalue()
                        preset_upload_identity = (
                            mapping_preset_upload.name,
                            hashlib.sha256(preset_upload_bytes).hexdigest(),
                        )
                        if st.session_state.get("tmc_mapping_preset_upload_identity") != preset_upload_identity:
                            if st.session_state.get("tmc_mapping_preset_ignored_identity") != preset_upload_identity:
                                st.session_state.pop("tmc_mapping_preset_ignored_identity", None)
                                loaded_preset = load_mapping_preset(preset_upload_bytes)
                                apply_result = apply_mapping_preset_to_detected_sheets(loaded_preset, detected_sheet_names)
                                default_mapping = apply_result.mapping
                                detected_scheme = detect_mapping_preset_scheme(loaded_preset)
                                _set_current_mapping_scheme(detected_scheme)
                                st.session_state["tmc_loaded_mapping_code_scheme"] = detected_scheme
                                st.session_state["mapping_table"] = default_mapping.to_dict("records")
                                st.session_state["mapping_editor_version"] = int(st.session_state.get("mapping_editor_version", 0) or 0) + 1
                                st.session_state["tmc_mapping_table_from_session"] = False
                                _set_mapping_source(MAPPING_SOURCE_MAPPING_PRESET)
                                st.session_state["tmc_mapping_preset_upload_identity"] = preset_upload_identity
                                st.session_state["tmc_mapping_preset_apply_info"] = {
                                    "matched": apply_result.matched_sheet_count,
                                    "missing": apply_result.missing_detected_sheet_count,
                                    "extra": apply_result.extra_preset_row_count,
                                    "scheme": detected_scheme,
                                }
                                st.session_state["tmc_mapping_preset_warnings"] = list(loaded_preset.warnings)
                    except (MappingPresetError, ValueError) as exc:
                        st.error(f"ไม่สามารถเปิด Mapping Preset ได้: {exc}")
                preset_info = st.session_state.get("tmc_mapping_preset_apply_info")
                if preset_info:
                    _render_alert("โหลด Mapping Preset สำเร็จ", "success")
                    _render_alert(
                        f"จับคู่ Sheet ได้ {preset_info.get('matched', 0)} รายการ; "
                        f"ยังต้องตรวจสอบ Sheet ที่พบใหม่ {preset_info.get('missing', 0)} รายการ; "
                        f"แถวจาก Preset ที่ไม่พบใน workbook นี้ {preset_info.get('extra', 0)} รายการ",
                        "info",
                    )
                for warning_message in st.session_state.get("tmc_mapping_preset_warnings", []):
                    _render_alert(warning_message, "warning")

                mapping_scheme = _current_mapping_scheme()
                mapping_editor_version = int(st.session_state.get("mapping_editor_version", 0) or 0)
                mapping_view = str(st.session_state.get("mapping_editor_view_mode") or "Basic")
                editor_key = _mapping_editor_key(mapping_view, mapping_editor_version)
                basic_widget_source = _mapping_source() in {MAPPING_SOURCE_DEFAULT_PREVIEW, MAPPING_SOURCE_USER_EDITOR}
                if basic_widget_source or mapping_view != "Basic":
                    default_mapping = _mapping_from_editor_widget_state(
                        default_mapping,
                        mapping_view,
                        mapping_scheme,
                        editor_key,
                        apply_basic_widgets=basic_widget_source,
                    )
                if _is_v2_scheme(mapping_scheme):
                    default_mapping = normalize_approach_movement_mapping(default_mapping)
                if basic_widget_source and _basic_mapping_widget_state_exists(
                    mapping_editor_version,
                    st.session_state,
                ):
                    st.session_state["mapping_table"] = default_mapping.to_dict("records")
                    st.session_state["tmc_mapping_table_from_session"] = False
                    if _mapping_source() == MAPPING_SOURCE_DEFAULT_PREVIEW:
                        _set_mapping_source(MAPPING_SOURCE_USER_EDITOR)
                scheme_validation_issues = validate_mapping_scheme(default_mapping, mapping_scheme)
                process_block_reason = _single_file_processing_block_reason(mapping_scheme)
                mapping_issues = _single_file_mapping_issues(detected_sheet_names, default_mapping, mapping_scheme)
                mapping_counts = _mapping_workspace_counts(default_mapping, detected_sheet_names)
                _render_mapping_scheme_status(mapping_scheme)
                for issue in scheme_validation_issues:
                    _render_alert(issue, "warning")
                _render_metric_strip(
                    [
                        ("ไฟล์สำรวจ", "โหลดแล้ว", "", uploaded_file.name, "พร้อม"),
                        ("Sheet ที่พบ", mapping_counts["detected_sheets"], "sheet", "ตรวจจาก workbook", "พร้อม"),
                        ("แถว Mapping", mapping_counts["rows"], "แถว", "ข้อมูลหลักสำหรับประมวลผล", "พร้อม" if mapping_counts["rows"] else "ต้องตรวจสอบ"),
                        ("Movement ที่ใช้", mapping_counts["included"], "แถว", f"ไม่รวม {mapping_counts['excluded']:,} แถว", "พร้อม" if mapping_counts["included"] else "ต้องตรวจสอบ"),
                        ("รวมหลาย source", mapping_counts["duplicate_movements"], "movement", "อนุญาตสำหรับ aggregation", "ข้อมูล" if mapping_counts["duplicate_movements"] else "พร้อม"),
                        ("สถานะ Mapping", "พร้อม" if mapping_issues.empty else "ต้องตรวจสอบ", "", "ประมวลผลได้" if mapping_issues.empty else f"{len(mapping_issues):,} รายการ", "พร้อม" if mapping_issues.empty else "ต้องตรวจสอบ"),
                    ],
                    columns=6,
                )

                action_col, readiness_col = st.columns([0.82, 1.18])
                process_requested = bool(st.session_state.get("tmc_single_file_process_requested"))
                with action_col:
                    run = st.button("ประมวลผลไฟล์ TMC", type="primary", disabled=bool(process_block_reason) or not mapping_issues.empty or bool(scheme_validation_issues), key="process_tmc_mapping_top")
                    process_status_placeholder = st.empty()
                    if process_requested:
                        process_status_placeholder.status("กำลังประมวลผลไฟล์ TMC...", expanded=True)
                with readiness_col:
                    if process_block_reason:
                        _render_alert(process_block_reason, "warning")
                    elif mapping_issues.empty and not scheme_validation_issues:
                        _render_alert("การกำหนดทิศทางพร้อมสำหรับประมวลผล", "success")
                    else:
                        _render_alert("กรุณาตรวจสอบ Mapping ก่อนประมวลผล", "warning")

                for warning_message in mapping_control_warnings(default_mapping, mapping_scheme):
                    _render_alert(_thai_mapping_control_warning(warning_message), "warning")
                movement_code_options = _movement_code_options_for_scheme(default_mapping, mapping_scheme)
                with st.expander("Sheet ทิศทางที่ตรวจพบ", expanded=False):
                    st.dataframe(preview_summary, width="stretch")
                _render_section_header("ตาราง Mapping", "ตารางนี้เป็นข้อมูลหลักสำหรับการจับคู่ movement")
                mapping_view = st.radio(
                    "มุมมองตาราง Mapping",
                    options=["Basic", "Advanced"],
                    horizontal=True,
                    key="mapping_editor_view_mode",
                    help="Basic แสดงคอลัมน์ที่ใช้บ่อย ส่วน Advanced แสดงคอลัมน์เสริมและเชิงเทคนิค",
                )
                editor_key = _mapping_editor_key(mapping_view, mapping_editor_version)
                editor_frame = _mapping_editor_frame(default_mapping, mapping_view, mapping_scheme)
                use_basic_controls = mapping_view == "Basic" and _mapping_source() in {
                    MAPPING_SOURCE_DEFAULT_PREVIEW,
                    MAPPING_SOURCE_USER_EDITOR,
                }
                if use_basic_controls:
                    edited_frame = _render_basic_mapping_editor(editor_frame, movement_code_options, mapping_editor_version)
                else:
                    mapping = st.data_editor(
                        editor_frame,
                        width="stretch",
                        num_rows="dynamic",
                        column_config=_mapping_editor_column_config(mapping_scheme, movement_code_options),
                        key=editor_key,
                    )
                    edited_frame = pd.DataFrame(mapping)
                    edited_frame = _apply_mapping_editor_widget_state(edited_frame, st.session_state.get(editor_key))
                editor_changed = _mapping_editor_changed(editor_frame, edited_frame)
                mapping = _merge_mapping_editor_result(default_mapping, edited_frame, mapping_scheme)
                if mapping_view == "Basic":
                    basic_summary = _basic_mapping_derived_summary(mapping, mapping_scheme)
                    st.table(basic_summary)
                    invalid_basic_rows = basic_summary[
                        basic_summary["Status"].astype(str).str.startswith("รหัสทิศทางไม่ถูกต้อง")
                    ]
                    if not invalid_basic_rows.empty:
                        _render_alert("รหัสทิศทางไม่ถูกต้อง ใช้รูปแบบเช่น NE, NS, NW, NU", "warning")
                if _mapping_rows_are_committed() or editor_changed:
                    st.session_state["mapping_table"] = mapping.to_dict("records")
                    st.session_state["tmc_mapping_table_from_session"] = False
                    if editor_changed and _mapping_source() == MAPPING_SOURCE_DEFAULT_PREVIEW:
                        _set_mapping_source(MAPPING_SOURCE_USER_EDITOR)
                mapping_counts = _mapping_workspace_counts(mapping, detected_sheet_names)

                aggregation_preview = _mapping_aggregation_preview(mapping)
                if not aggregation_preview.empty:
                    _render_alert("พบ movement ที่รวมจากหลาย source stream", "info")
                    with st.expander("Audit: movement ที่รวมจากหลาย source stream", expanded=False):
                        st.dataframe(aggregation_preview, width="stretch", hide_index=True)
                        for aggregation_message in movement_aggregation_messages(mapping):
                            st.caption(_thai_aggregation_message(aggregation_message))
    
                mapping_scheme = _current_mapping_scheme()
                scheme_validation_issues = validate_mapping_scheme(mapping, mapping_scheme)
                process_block_reason = _single_file_processing_block_reason(mapping_scheme)
                mapping_issues = _single_file_mapping_issues(detected_sheet_names, mapping, mapping_scheme)
                if process_block_reason:
                    _render_alert(process_block_reason, "warning")
                elif _is_v2_scheme(mapping_scheme) and not scheme_validation_issues:
                    _render_alert("Mapping approach_movement พร้อมใช้งานสำหรับ single-file workflow", "success")
                elif mapping_issues.empty and not scheme_validation_issues:
                    _render_alert("Mapping พร้อมใช้งาน", "success")
                else:
                    _render_alert("มีรายการที่ต้องตรวจสอบก่อนประมวลผล", "warning")
                    with st.expander("รายการที่ต้องตรวจสอบ", expanded=True):
                        st.dataframe(_mapping_issue_display(mapping_issues), width="stretch")
                if mapping_counts["blank_source_stream"]:
                    _render_alert(f"พบ source_stream ว่าง {mapping_counts['blank_source_stream']:,} แถว ระบบจะใช้ค่าเริ่มต้นตามพฤติกรรมเดิมเมื่อรวมข้อมูล", "info")
                if mapping_counts["excluded"]:
                    _render_alert(f"มีแถวที่ไม่แสดงในรายงาน {mapping_counts['excluded']:,} แถว", "info")

                _sync_single_workflow_from_state(
                    source_bytes=file_bytes if uploaded_file is not None else None,
                    source_file_name=uploaded_file.name if uploaded_file is not None else None,
                    export_mode=export_mode,
                )

                if run and mapping_issues.empty and not process_block_reason and not scheme_validation_issues:
                    st.session_state["tmc_single_file_process_requested"] = True
                    st.rerun()

                if process_requested and mapping_issues.empty and not process_block_reason and not scheme_validation_issues:
                    st.session_state.pop("tmc_single_file_process_requested", None)
                    try:
                        raw_sheets = {name: parsed.data for name, parsed in parsed_details.items()}
                        with process_status_placeholder.status("กำลังประมวลผลไฟล์ TMC...", expanded=True) as processing_status:
                            try:
                                result = _process_single_file_for_ui(
                                    raw_sheets=raw_sheets,
                                    mapping=mapping,
                                    setup=setup,
                                    detected_sheets=detected_sheet_names,
                                    peak_mode=peak_mode,
                                    peak_windows=peak_windows,
                                    pce_factors=selected_pce_factors,
                                )
                            except Exception:
                                processing_status.update(label="กำลังประมวลผลไฟล์ TMC...", state="error", expanded=True)
                                raise
                            processing_status.update(label="กำลังประมวลผลไฟล์ TMC...", state="complete", expanded=False)
                    except Exception as exc:  # pragma: no cover - UI guardrail
                        st.error(f"ประมวลผลไม่สำเร็จ: {exc}")
                    else:
                        st.session_state["tmc_processed"] = {
                            "result": result,
                            "mapping": mapping,
                            "setup": setup,
                            "peak_windows": peak_windows,
                            "pce_factors": selected_pce_factors,
                        }
                        st.session_state.pop("tmc_output", None)
                        st.session_state.pop("tmc_pce_results_stale", None)
                        processed_revisions = _single_workflow_revisions(
                            source_bytes=file_bytes if uploaded_file is not None else None,
                            source_file_name=uploaded_file.name if uploaded_file is not None else None,
                            mapping=mapping,
                            pce_factors=selected_pce_factors,
                            peak_mode=peak_mode,
                            peak_windows=peak_windows,
                            movement_code_scheme=_current_mapping_scheme(),
                            setup=setup,
                            export_mode=export_mode,
                            confirmed_peaks=_confirmed_peaks_from_state(),
                            analysis_present=True,
                        )
                        _record_workflow_state(
                            WORKFLOW_SINGLE_MODE,
                            processed_revisions,
                            WorkflowReadiness(
                                source=uploaded_file is not None,
                                mapping=bool(mapping.to_dict("records")),
                                analysis=True,
                                review=bool(_single_effective_peak_state().get("ready")),
                                export=False,
                            ),
                        )
                        set_active_tab("ตรวจ Peak")
                        _flash_and_rerun("ประมวลผลเสร็จแล้ว กรุณาตรวจสอบช่วงเร่งด่วนในแท็บ “ตรวจ Peak”")
                elif process_requested:
                    st.session_state.pop("tmc_single_file_process_requested", None)
                    process_status_placeholder.empty()

    if not is_single_file_mode:
        batch_export_mode = st.session_state.get("tmc_batch_export_mode", BATCH_SAFE_PNG_EXPORT_LABEL)
        uploaded_ready = bool(batch_uploads)
        mapping_ready = loaded_batch_preset is not None
        pce_ready = bool(selected_pce_factors)
        batch_ready = batch_inputs_ready(
            uploaded_workbook_count=len(batch_uploads or []),
            mapping_available=mapping_ready,
            pce_factors_ready=pce_ready,
            movement_code_scheme=batch_mapping_scheme,
        )
        batch_process_block_reason = batch_processing_block_reason(batch_mapping_scheme)
        batch_signature = _batch_analysis_signature(
            uploads_signature=_batch_upload_signature(batch_uploads),
            preset_signature=batch_preset_signature,
            pce_factors=selected_pce_factors,
            peak_mode=peak_mode,
            peak_windows=peak_windows,
        )
        batch_stale = _mark_batch_stale_if_inputs_changed(batch_signature)
        batch_analysis = st.session_state.get("tmc_batch_analysis_result")
        batch_result = st.session_state.get("tmc_batch_export_result")
        batch_export_signature = _batch_export_signature(
            metadata_rows=st.session_state.get("tmc_batch_file_metadata_table") or [],
            shared_setup=setup,
            export_mode=batch_export_mode,
            confirmed_peaks=st.session_state.get("tmc_batch_confirmed_peaks") or {},
        )
        batch_export_stale = _mark_batch_export_stale_if_inputs_changed(batch_export_signature)

        if active_tab == "ตั้งค่า":
            _render_section_header("ตั้งค่า Batch", "ใช้สำหรับจุดสำรวจเดียวกันหลายวัน โดยใช้ Mapping Preset เดียวกันทุกไฟล์")
            if batch_stale:
                _render_alert("ข้อมูล Batch มีการเปลี่ยนแปลง กรุณาวิเคราะห์ Batch ใหม่", "warning")

            batch_left, batch_right = st.columns([1.1, 1])
            with batch_left:
                with st.container(border=True):
                    _render_section_header("ขอบเขต Batch", "ค่ากลางที่จะใช้ร่วมกันทุกไฟล์")
                    _render_readiness_checklist(
                        [
                            ("ใช้ Mapping Preset เดียวกันทุกไฟล์", mapping_ready, "เปิดไฟล์ Preset ใน sidebar"),
                            ("ใช้ค่า PCE ชุดเดียวกัน", pce_ready, ""),
                            ("ตรวจ Peak แยกแต่ละไฟล์", bool(batch_analysis), "อยู่ในแท็บตรวจ Peak"),
                            ("ไม่รวม raw Excel ใน ZIP", True, "แพ็กเกจส่งออกมีเฉพาะรายงานและไฟล์ประกอบ"),
                        ]
                    )
                    setup_cols = st.columns(3)
                    setup_cols[0].text_input("ชื่อจุดนับ / TMC title", key="tmc_title_input")
                    setup_cols[1].text_input("จุดสำรวจ", key="survey_point_input")
                    setup_cols[2].text_input("ช่วงเวลาสำรวจ", key="survey_period_input")
                    direction_cols = st.columns(4)
                    direction_cols[0].text_input("ป้ายปลายทางด้านเหนือ", key="north_label_input")
                    direction_cols[1].text_input("ป้ายปลายทางด้านใต้", key="south_label_input")
                    direction_cols[2].text_input("ป้ายปลายทางด้านตะวันออก", key="east_label_input")
                    direction_cols[3].text_input("ป้ายปลายทางด้านตะวันตก", key="west_label_input")
                    road_cols = st.columns(4)
                    road_cols[0].text_input("ชื่อถนนด้านเหนือ", key="north_road_input")
                    road_cols[1].text_input("ชื่อถนนด้านใต้", key="south_road_input")
                    road_cols[2].text_input("ชื่อถนนด้านตะวันออก", key="east_road_input")
                    road_cols[3].text_input("ชื่อถนนด้านตะวันตก", key="west_road_input")

                with st.container(border=True):
                    _render_section_header("ค่า PCE ร่วม", "ค่า PCE ชุดนี้จะใช้กับทุกไฟล์ใน Batch")
                    selected_pce_factors = _render_pce_factor_editor()
                    _sync_workflow_after_pce_editor(
                        is_single_file_mode=False,
                        export_mode=str(st.session_state.get("tmc_batch_export_mode") or BATCH_SAFE_PNG_EXPORT_LABEL),
                        batch_uploads=batch_uploads,
                        mapping_preset=loaded_batch_preset,
                        movement_code_scheme=batch_mapping_scheme,
                        metadata_rows=batch_metadata_rows,
                    )
                    has_overrides, override_text = _pce_override_summary(selected_pce_factors)
                    _render_status_chip("มีค่า PCE ที่แก้ไขเอง" if has_overrides else "ใช้ค่า PCE เริ่มต้น", "warning" if has_overrides else "success")
                    if has_overrides:
                        st.caption(override_text)

            with batch_right:
                with st.container(border=True):
                    _render_section_header("ไฟล์ที่อัปโหลด", "ตรวจจำนวนไฟล์และชื่อไฟล์ก่อนวิเคราะห์ Batch")
                    if not batch_metadata_rows:
                        _render_action_hint("อัปโหลดไฟล์ TMC Excel หลายไฟล์ในแถบด้านซ้าย")
                    else:
                        _render_metric_strip(
                            [
                                ("จำนวนไฟล์", f"{len(batch_metadata_rows):,}", "ไฟล์", "พร้อมตั้งค่ารายไฟล์"),
                            ],
                            columns=1,
                        )
                        st.dataframe(pd.DataFrame({"file_name": [row["file_name"] for row in batch_metadata_rows]}), width="stretch")

                with st.container(border=True):
                    _render_section_header("ข้อมูลรายไฟล์", "survey_date_text และ output_stem จะใช้ในรายงานและ ZIP")
                    if not batch_metadata_rows:
                        _render_empty_state("ยังไม่มีข้อมูลรายไฟล์", "อัปโหลดไฟล์ Batch เพื่อสร้างตารางตั้งค่า")
                    else:
                        metadata_version = int(st.session_state.get("tmc_batch_file_metadata_editor_version", 0) or 0)
                        edited_metadata = st.data_editor(
                            pd.DataFrame(batch_metadata_rows),
                            key=f"tmc_batch_file_metadata_editor_{metadata_version}",
                            hide_index=True,
                            width="stretch",
                            disabled=["file_name"],
                            column_config={
                                "file_name": st.column_config.TextColumn("ชื่อไฟล์ต้นทาง"),
                                "survey_date_text": st.column_config.TextColumn("วันที่สำรวจ"),
                                "output_stem": st.column_config.TextColumn("ชื่อไฟล์ส่งออก"),
                                "notes": st.column_config.TextColumn("หมายเหตุ"),
                            },
                        )
                        cleaned_metadata = []
                        for row in edited_metadata.to_dict("records"):
                            cleaned_metadata.append(
                                {
                                    "file_name": Path(str(row.get("file_name", ""))).name,
                                    "survey_date_text": str(row.get("survey_date_text", "") or ""),
                                    "output_stem": safe_output_stem(str(row.get("output_stem", "") or row.get("file_name", ""))),
                                    "notes": str(row.get("notes", "") or ""),
                                }
                            )
                        if cleaned_metadata != st.session_state.get("tmc_batch_file_metadata_table"):
                            st.session_state["tmc_batch_file_metadata_table"] = cleaned_metadata
                            _sync_batch_analysis_metadata_from_state()
                            _mark_batch_export_stale_now()
                            _sync_batch_workflow_from_state(
                                batch_uploads=batch_uploads,
                                mapping_preset=loaded_batch_preset,
                                movement_code_scheme=batch_mapping_scheme,
                                metadata_rows=cleaned_metadata,
                                export_mode=batch_export_mode,
                            )

        if active_tab == "กำหนดทิศทาง":
            _render_section_header(
                "กำหนดทิศทาง Batch",
                "ใช้ Mapping Preset เดียวกันสำหรับไฟล์สำรวจหลายวันของจุดเดียวกัน",
            )
            preset_rows = _mapping_preset_rows_frame(loaded_batch_preset)
            preset_code_column = "output_movement_code" if "output_movement_code" in preset_rows else "movement_code"
            preset_included = 0
            preset_duplicate_count = 0
            if not preset_rows.empty:
                preset_include = preset_rows["include_in_report"].fillna(True).astype(bool) if "include_in_report" in preset_rows else pd.Series(True, index=preset_rows.index)
                preset_codes = preset_rows[preset_code_column].fillna("").astype(str).str.strip() if preset_code_column in preset_rows else pd.Series("", index=preset_rows.index)
                preset_included = int((preset_include & (preset_codes != "")).sum())
                preset_duplicate_count = int((preset_codes[preset_include & (preset_codes != "")].value_counts() > 1).sum())

            _render_metric_strip(
                [
                    ("Mapping Preset", "พร้อมใช้งาน" if loaded_batch_preset else "ยังไม่พร้อม", "", batch_preset_name if loaded_batch_preset else "เปิดไฟล์ Preset ที่ sidebar", "พร้อมใช้งาน" if loaded_batch_preset else "ต้องตรวจสอบ"),
                    ("แถว Mapping", len(preset_rows), "แถว", "shared preset", "พร้อม" if loaded_batch_preset else "ต้องตรวจสอบ"),
                    ("Movement ที่ใช้", preset_included, "แถว", "include_in_report", "พร้อม" if preset_included else "ต้องตรวจสอบ"),
                    ("รวมหลาย source", preset_duplicate_count, "movement", "อนุญาตสำหรับ aggregation", "ข้อมูล" if preset_duplicate_count else "พร้อม"),
                ],
                columns=4,
            )
            if loaded_batch_preset:
                _render_mapping_scheme_status(batch_mapping_scheme)

            if not batch_uploads:
                _render_action_hint("อัปโหลดไฟล์ TMC หลายไฟล์ที่แถบด้านซ้ายก่อนตรวจ Mapping")
            if not loaded_batch_preset:
                _render_action_hint("เปิด Mapping Preset เพื่อใช้กับไฟล์ Batch")

            _render_section_header("ความพร้อม Batch", "ตรวจรายการจำเป็นก่อนวิเคราะห์ในแท็บ ตรวจ Peak")
            _render_readiness_checklist(
                [
                    ("อัปโหลดไฟล์ Batch", uploaded_ready, f"{len(batch_uploads or []):,} ไฟล์" if uploaded_ready else "ยังไม่มีไฟล์"),
                    ("Mapping Preset พร้อมใช้งาน", mapping_ready, batch_preset_name if mapping_ready else "ยังไม่เปิด Preset"),
                    ("ค่า PCE พร้อมใช้งาน", pce_ready, "พร้อมใช้งาน" if pce_ready else "ตรวจค่า PCE ในแท็บตั้งค่า"),
                    ("Metadata รายไฟล์พร้อมใช้งาน", bool(batch_metadata_rows), f"{len(batch_metadata_rows):,} แถว" if batch_metadata_rows else "ตั้งค่ารายไฟล์ในแท็บตั้งค่า"),
                ]
            )
            if mapping_ready:
                if batch_process_block_reason:
                    _render_alert(batch_process_block_reason, "warning")
                else:
                    _render_action_hint("เมื่อ Mapping พร้อมใช้งานแล้ว ไปที่แท็บ ตรวจ Peak เพื่อวิเคราะห์ Batch")

            with st.expander("สถานะ Sheet matching รายไฟล์", expanded=False):
                _render_action_hint("ตรวจว่า Sheet ในแต่ละไฟล์ตรงกับ Mapping Preset แค่ไหน")
                if not batch_uploads or not loaded_batch_preset:
                    _render_empty_state("ยังตรวจ Sheet matching ไม่ได้", "อัปโหลดไฟล์ Batch และเปิด Mapping Preset ก่อน")
                else:
                    status_rows = []
                    for file in batch_uploads:
                        try:
                            detected = detect_raw_direction_sheet_names(BytesIO(file.getvalue()))
                            apply_result = apply_mapping_preset_to_detected_sheets(loaded_batch_preset, detected)
                            missing_sheets = ", ".join(apply_result.missing_detected_sheets)
                            extra_sheets = ", ".join(apply_result.extra_preset_sheets)
                            status_rows.append(
                                {
                                    "file_name": Path(file.name).name,
                                    "detected_sheets": len(detected),
                                    "matched_sheets": apply_result.matched_sheet_count,
                                    "missing_detected_sheets": apply_result.missing_detected_sheet_count,
                                    "preset_rows_not_found": apply_result.extra_preset_row_count,
                                    "mapping_status": "พร้อม" if apply_result.missing_detected_sheet_count == 0 else "ต้องตรวจสอบ",
                                    "notes": missing_sheets or extra_sheets or "",
                                }
                            )
                        except Exception as exc:
                            status_rows.append(
                                {
                                    "file_name": Path(file.name).name,
                                    "detected_sheets": 0,
                                    "matched_sheets": 0,
                                    "missing_detected_sheets": 0,
                                    "preset_rows_not_found": 0,
                                    "mapping_status": "อ่านไฟล์ไม่สำเร็จ",
                                    "notes": str(exc),
                                }
                            )
                    st.dataframe(pd.DataFrame(status_rows), width="stretch", hide_index=True)

            if loaded_batch_preset and not preset_rows.empty:
                with st.expander("ตัวอย่าง Mapping Preset", expanded=False):
                    preview_columns = [
                        column
                        for column in [
                            "raw_sheet",
                            "raw_direction",
                            "source_stream",
                            "raw_movement_label",
                            "output_movement_code",
                            "include_in_report",
                            "include_in_peak",
                            "aggregation_method",
                        ]
                        if column in preset_rows.columns
                    ]
                    st.dataframe(preset_rows[preview_columns].head(50) if preview_columns else preset_rows.head(50), width="stretch", hide_index=True)

        if active_tab == "ตรวจ Peak":
            _render_section_header("ตรวจ Peak รายไฟล์", "ตรวจกราฟและยืนยัน AM/PM Peak แยกตามไฟล์ ก่อนสร้าง Batch ZIP")
            if batch_stale:
                _render_alert("ข้อมูล Batch มีการเปลี่ยนแปลง กรุณาวิเคราะห์ Batch ใหม่", "warning")
            if not batch_ready:
                _render_readiness_checklist(
                    [
                        ("อัปโหลดไฟล์ Batch", uploaded_ready, f"{len(batch_uploads or []):,} ไฟล์" if uploaded_ready else "ยังไม่ได้อัปโหลดไฟล์ TMC Excel"),
                        ("Mapping Preset", mapping_ready, "พร้อมใช้งาน" if mapping_ready else "เปิด Mapping Preset ใน sidebar"),
                        ("ค่า PCE", pce_ready, "พร้อมใช้งาน" if pce_ready else "ตรวจสอบค่า PCE ในแท็บตั้งค่า"),
                    ]
                )
                if batch_process_block_reason:
                    _render_alert(batch_process_block_reason, "warning")
                else:
                    _render_action_hint("เตรียมไฟล์ Mapping Preset และค่า PCE ให้พร้อมก่อนวิเคราะห์ Batch")
            elif not batch_analysis or batch_stale:
                _render_action_hint("วิเคราะห์ Batch เพื่อสร้างรายการตรวจ Peak รายไฟล์")
            analyze_batch = st.button("วิเคราะห์ Batch", type="primary", disabled=not batch_ready, key="analyze_batch_processing")
            if analyze_batch:
                items = _batch_items_from_uploads(batch_uploads)
                with st.spinner("กำลังวิเคราะห์ Batch..."):
                    batch_analysis = analyze_batch_files(
                        items,
                        mapping_preset=loaded_batch_preset,
                        setup={**setup, "movement_code_scheme": batch_mapping_scheme},
                        pce_factors=selected_pce_factors,
                        peak_mode=peak_mode,
                        peak_windows=peak_windows,
                        mapping_preset_name=batch_preset_name,
                    )
                st.session_state["tmc_batch_analysis_result"] = batch_analysis
                st.session_state["tmc_batch_preset_name"] = batch_preset_name
                st.session_state["tmc_batch_input_signature"] = st.session_state.get("tmc_batch_current_input_signature")
                st.session_state["tmc_batch_stale"] = False
                st.session_state["tmc_batch_export_stale"] = False
                st.session_state.pop("tmc_batch_export_result", None)
                st.session_state["tmc_batch_review_version"] = int(st.session_state.get("tmc_batch_review_version", 0) or 0) + 1
                st.session_state["tmc_batch_confirmed_peaks"] = {
                    item.folder_name: {"AM": item.confirmed_AM_peak, "PM": item.confirmed_PM_peak}
                    for item in batch_analysis.successful_items
                }
                analyzed_batch_revisions = _batch_workflow_revisions(
                    uploads=batch_uploads,
                    mapping_preset=loaded_batch_preset,
                    pce_factors=selected_pce_factors,
                    peak_mode=peak_mode,
                    peak_windows=peak_windows,
                    movement_code_scheme=batch_mapping_scheme,
                    metadata_rows=st.session_state.get("tmc_batch_file_metadata_table") or [],
                    setup=setup,
                    export_mode=batch_export_mode,
                    confirmed_peaks=_stable_batch_confirmed_peaks(
                        st.session_state.get("tmc_batch_confirmed_peaks") or {},
                        batch_analysis,
                    ),
                    analysis_present=True,
                )
                _record_workflow_state(
                    WORKFLOW_BATCH_MODE,
                    analyzed_batch_revisions,
                    WorkflowReadiness(
                        source=bool(batch_uploads),
                        mapping=loaded_batch_preset is not None,
                        analysis=True,
                        review=bool(batch_analysis.successful_items)
                        and all(
                            item.confirmed_AM_peak and item.confirmed_PM_peak
                            for item in batch_analysis.successful_items
                        ),
                        export=False,
                    ),
                )
                set_active_tab("ตรวจ Peak")
                _flash_and_rerun("วิเคราะห์ Batch เสร็จแล้ว กรุณาตรวจสอบช่วงเร่งด่วนก่อนสร้าง ZIP")

            if batch_analysis:
                batch_confirmed_peaks = st.session_state.setdefault("tmc_batch_confirmed_peaks", {})
                batch_review_version = int(st.session_state.get("tmc_batch_review_version", 0) or 0)
                for item in batch_analysis.successful_items:
                    stored = batch_confirmed_peaks.get(item.folder_name, {})
                    item.confirmed_AM_peak = stored.get("AM", item.confirmed_AM_peak)
                    item.confirmed_PM_peak = stored.get("PM", item.confirmed_PM_peak)
                successful_items = batch_analysis.successful_items
                confirmed_count = sum(1 for item in successful_items if item.confirmed_AM_peak and item.confirmed_PM_peak)
                successful_count = len(successful_items)
                status_items = [
                    ("ไฟล์ทั้งหมด", f"{len(batch_analysis.items):,}", "ไฟล์", ""),
                    ("กำหนด Peak", f"{confirmed_count:,}/{successful_count:,}", "ไฟล์", "ไฟล์ที่วิเคราะห์สำเร็จ"),
                    ("ไฟล์ไม่สำเร็จ", f"{sum(1 for item in batch_analysis.items if item.status == 'failed'):,}", "ไฟล์", "ไม่ต้องกำหนด Peak"),
                ]
                _render_metric_strip(status_items, columns=3)
                if successful_count and confirmed_count == successful_count:
                    _render_alert("กำหนด Peak ครบแล้ว พร้อมส่งออก Batch", "success")
                elif successful_count:
                    _render_alert("ยังมีไฟล์ที่ต้องกำหนด Peak", "warning")
                st.dataframe(_batch_peak_review_display_frame(batch_analysis), width="stretch", hide_index=True)

                if successful_items:
                    review_labels = {f"{item.file_name} ({item.survey_date_text or 'no date'})": item.folder_name for item in successful_items}
                    selected_folder = st.session_state.get("tmc_batch_selected_review_file") or successful_items[0].folder_name
                    label_by_folder = {folder: label for label, folder in review_labels.items()}
                    selected_label = label_by_folder.get(selected_folder, next(iter(review_labels)))
                    selected_review_label = st.selectbox(
                        "เลือกไฟล์สำหรับตรวจกราฟ",
                        options=list(review_labels),
                        index=list(review_labels).index(selected_label),
                        key="tmc_batch_selected_review_label",
                    )
                    selected_folder = review_labels[selected_review_label]
                    st.session_state["tmc_batch_selected_review_file"] = selected_folder
                    selected_item = next(item for item in successful_items if item.folder_name == selected_folder)
                    preview = batch_selected_file_preview(selected_item)
                    with st.container(border=True):
                        _render_section_header("ไฟล์ที่เลือก", f"{preview['file_name']} · {preview['output_stem']}")
                        _render_metric_strip(
                            [
                                ("วันที่สำรวจ", preview["survey_date_text"] or "-", "", ""),
                                ("สถานะ", preview["status"] or "-", "", "", preview["status"] or "pending"),
                                ("จำนวนรถรวม", format_count(preview["total_vehicles"]), "คัน", ""),
                                ("PCU รวม", format_pcu(preview["total_PCU"]), "PCU", ""),
                            ],
                            columns=4,
                        )
                        _render_metric_strip(
                            [
                                ("QC error", format_count(preview["QC_errors"]), "", ""),
                                ("QC warning", format_count(preview["QC_warnings"]), "", ""),
                                ("QC info", format_count(preview["QC_info"]), "", ""),
                            ],
                            columns=3,
                        )
                    _render_section_header("กราฟ PCU รายชั่วโมง", "ใช้ตรวจรูปแบบปริมาณจราจรก่อนกำหนดช่วง Peak")
                    _render_hourly_pcu_line_chart(
                        selected_item.hourly_movement_pcu,
                        preview["confirmed_AM_peak"] or preview["suggested_AM_peak"] or "",
                        preview["confirmed_PM_peak"] or preview["suggested_PM_peak"] or "",
                    )
                    option_labels = list(dict.fromkeys(selected_item.hourly_period_options or [selected_item.suggested_AM_peak, selected_item.suggested_PM_peak]))
                    option_labels = [value for value in option_labels if value]
                    _render_section_header("กำหนด Peak ของไฟล์นี้", "ระบบจะใช้ช่วง Peak ที่กำหนดในหน้านี้สำหรับรายงานของไฟล์นี้")
                    peak_cols = st.columns(2)
                    if option_labels:
                        stored = batch_confirmed_peaks.setdefault(selected_item.folder_name, {"AM": selected_item.confirmed_AM_peak, "PM": selected_item.confirmed_PM_peak})
                        am_default = stored.get("AM") or selected_item.confirmed_AM_peak
                        pm_default = stored.get("PM") or selected_item.confirmed_PM_peak
                        for value in [am_default, pm_default]:
                            if value and value not in option_labels:
                                option_labels.insert(0, value)
                        option_labels = list(dict.fromkeys(option_labels))
                        with peak_cols[0]:
                            _render_peak_card("AM Peak · ระบบตรวจจับอัตโนมัติ", preview["suggested_AM_peak"] or "", "", "auto_suggested")
                            selected_am = st.selectbox(
                                "ช่วงที่กำหนด AM",
                                options=option_labels,
                                index=option_labels.index(am_default) if am_default in option_labels else 0,
                                key=f"batch_review_am_{batch_review_version}_{selected_item.folder_name}",
                            )
                            _render_status_chip("กำหนดแล้ว" if selected_am else "รอตรวจสอบ", "success" if selected_am else "warning")
                            _render_action_hint("ใช้ช่วงนี้เป็นค่าหลักสำหรับรายงาน")
                        with peak_cols[1]:
                            _render_peak_card("PM Peak · ระบบตรวจจับอัตโนมัติ", preview["suggested_PM_peak"] or "", "", "auto_suggested")
                            selected_pm = st.selectbox(
                                "ช่วงที่กำหนด PM",
                                options=option_labels,
                                index=option_labels.index(pm_default) if pm_default in option_labels else 0,
                                key=f"batch_review_pm_{batch_review_version}_{selected_item.folder_name}",
                            )
                            _render_status_chip("กำหนดแล้ว" if selected_pm else "รอตรวจสอบ", "success" if selected_pm else "warning")
                            _render_action_hint("ใช้ช่วงนี้เป็นค่าหลักสำหรับรายงาน")
                        review_changed = stored.get("AM") != selected_am or stored.get("PM") != selected_pm
                        batch_confirmed_peaks[selected_item.folder_name] = {"AM": selected_am, "PM": selected_pm}
                        selected_item.confirmed_AM_peak = selected_am
                        selected_item.confirmed_PM_peak = selected_pm
                        review_transition = _sync_batch_workflow_from_state(
                            batch_uploads=batch_uploads,
                            mapping_preset=loaded_batch_preset,
                            movement_code_scheme=batch_mapping_scheme,
                            metadata_rows=st.session_state.get("tmc_batch_file_metadata_table") or [],
                            export_mode=batch_export_mode,
                        )
                        if review_changed and not review_transition.export_invalidated:
                            _mark_batch_export_stale_now()
                    else:
                        _render_alert("ไม่มีช่วงเวลารายชั่วโมงสำหรับกำหนด Peak ของไฟล์นี้", "warning")
                if batch_analysis.has_failures:
                    _render_alert("บางไฟล์วิเคราะห์ไม่สำเร็จ ไฟล์เหล่านี้ยังแสดงในตารางและไม่ต้องกำหนด Peak", "warning")
            else:
                _render_empty_state("กรุณาวิเคราะห์ Batch ก่อนตรวจสอบ Peak", "กด วิเคราะห์ Batch เพื่อสร้างตารางตรวจ Peak รายไฟล์")

        if active_tab == "ส่งออก":
            _render_section_header("ส่งออก Batch", "สร้าง Batch ZIP พร้อมรายงานรายไฟล์และ batch_summary.xlsx")
            batch_export_options = _batch_export_mode_options(excel_com_status, batch_mapping_scheme)
            previous_batch_export_mode = st.session_state.get("tmc_batch_export_mode", batch_export_mode)
            selected_batch_export_mode = st.radio(
                "โหมดส่งออก Batch",
                options=batch_export_options,
                index=batch_export_options.index(
                    _coerce_export_mode(
                        previous_batch_export_mode,
                        batch_export_options,
                        _default_batch_export_mode(excel_com_status, batch_mapping_scheme),
                    )
                ),
                key="tmc_batch_export_mode_control",
                horizontal=True,
                help="Excel Template Mode รักษา Native Chart และรูปแบบ Excel Template เมื่อ Excel COM พร้อมใช้งาน. Safe PNG Export Mode ใช้กราฟ PNG แบบคงที่.",
            )
            if apply_batch_export_mode_change(selected_batch_export_mode, previous_batch_export_mode):
                st.rerun()
            batch_export_mode = selected_batch_export_mode
            batch_analysis = st.session_state.get("tmc_batch_analysis_result")
            batch_result = st.session_state.get("tmc_batch_export_result")
            no_successful_files = not batch_analysis or not batch_analysis.successful_items
            peaks_ready = bool(batch_analysis and reviewed_peak_values_complete(batch_analysis))
            output_stems_valid = all(str(row.get("output_stem", "")).strip() for row in st.session_state.get("tmc_batch_file_metadata_table") or [])
            v2_batch_template_mode_blocked = (
                _is_v2_scheme(batch_mapping_scheme)
                and batch_export_mode.startswith(BATCH_EXCEL_TEMPLATE_EXPORT_MODE)
            )
            export_mode_ready = bool(
                not v2_batch_template_mode_blocked
                and (
                    batch_export_mode.startswith(BATCH_SAFE_PNG_EXPORT_MODE)
                    or excel_com_status.available
                    or not batch_export_mode.startswith(BATCH_EXCEL_TEMPLATE_EXPORT_MODE)
                )
            )
            if batch_stale:
                _render_alert("ข้อมูล Batch มีการเปลี่ยนแปลง กรุณาวิเคราะห์ Batch ใหม่", "warning")
            elif batch_export_stale:
                _render_alert("ข้อมูลส่งออกมีการเปลี่ยนแปลง กรุณาสร้าง Batch ZIP ใหม่", "warning")
            block_reason = batch_zip_generation_block_reason(
                has_successful_files=not no_successful_files,
                peaks_ready=peaks_ready,
                batch_stale=batch_stale,
            )
            generate_disabled = bool(block_reason) or not output_stems_valid or not export_mode_ready
            if not output_stems_valid and not block_reason:
                block_reason = "ตรวจสอบ output_stem ในแท็บตั้งค่า Batch ก่อนสร้าง ZIP"
            if v2_batch_template_mode_blocked and not block_reason:
                block_reason = BATCH_V2_TEMPLATE_MODE_UNSUPPORTED_TH
            if not export_mode_ready and not block_reason:
                block_reason = "โหมดส่งออกยังไม่พร้อม"
            _render_action_hint(block_reason or "พร้อมสร้าง Batch ZIP")
            generate_batch = st.button("Generate Batch ZIP", type="primary", disabled=generate_disabled, key="generate_batch_zip")
            if generate_batch and batch_analysis:
                set_active_tab("ส่งออก")
                block_reason = batch_zip_generation_block_reason(
                    has_successful_files=bool(batch_analysis.successful_items),
                    peaks_ready=reviewed_peak_values_complete(batch_analysis),
                    batch_stale=bool(st.session_state.get("tmc_batch_stale")),
                )
                if block_reason:
                    _render_alert(block_reason, "warning")
                    st.stop()
                with st.spinner("กำลังสร้าง Batch ZIP..."):
                    batch_result = generate_batch_zip_from_reviewed_peaks(
                        batch_analysis,
                        setup=setup,
                        pce_factors=selected_pce_factors,
                        peak_mode=peak_mode,
                        peak_windows=peak_windows,
                        export_mode=batch_export_mode,
                        use_template_report_layout=_use_template_layout_for_export(batch_export_mode),
                        use_excel_com_native_charts=_use_excel_native_charts_for_export(batch_export_mode, excel_com_status),
                )
                st.session_state["tmc_batch_export_result"] = batch_result
                st.session_state["tmc_batch_export_stale"] = False
                st.session_state["tmc_batch_export_signature"] = _batch_export_signature(
                    metadata_rows=st.session_state.get("tmc_batch_file_metadata_table") or [],
                    shared_setup=setup,
                    export_mode=batch_export_mode,
                    confirmed_peaks=st.session_state.get("tmc_batch_confirmed_peaks") or {},
                )
                set_active_tab("ส่งออก")
                _flash_and_rerun("สร้าง Batch ZIP เสร็จแล้ว")

            batch_export_left, batch_export_right = st.columns([0.95, 1.05])
            with batch_export_left:
                with st.container(border=True):
                    _render_section_header("โหมดส่งออก", "สถานะโหมดที่เลือกสำหรับ Batch")
                    _render_status_chip(batch_export_mode, "success" if export_mode_ready else "warning")
                    if batch_export_mode.startswith(BATCH_EXCEL_TEMPLATE_EXPORT_MODE):
                        _render_alert(
                            "Excel Template Mode: เหมาะสำหรับรายงานฉบับใช้งานจริง รักษา Native Chart และรูปแบบ Excel Template เมื่อ Excel COM พร้อมใช้งาน",
                            "info",
                        )
                    else:
                        _render_alert(
                            "Safe PNG Export Mode: โหมดสำรอง เหมาะสำหรับตรวจร่างหรือกรณี Excel COM ใช้งานไม่ได้",
                            "info",
                        )
                    if batch_export_mode.startswith(BATCH_EXCEL_TEMPLATE_EXPORT_MODE) and len(batch_uploads or []) > 10:
                        _render_alert("มีไฟล์มากกว่า 10 ไฟล์ การสร้างรายงานด้วย Excel Template Mode อาจใช้เวลานานขึ้น", "warning")

            with batch_export_right:
                with st.container(border=True):
                    _render_section_header("ความพร้อม Batch", "ตรวจเงื่อนไขก่อนสร้าง Batch ZIP")
                    _render_readiness_checklist(
                        [
                            ("อัปโหลดไฟล์แล้ว", uploaded_ready, f"{len(batch_uploads or []):,} ไฟล์" if uploaded_ready else "ยังไม่มีไฟล์"),
                            ("Mapping Preset พร้อม", mapping_ready, "ใช้ Preset เดียวกันทุกไฟล์"),
                            ("Batch Analysis วิเคราะห์แล้ว", bool(batch_analysis and not batch_stale), ""),
                            ("กำหนด Peak ของไฟล์ที่สำเร็จแล้ว", peaks_ready, ""),
                            ("output_stem ใช้งานได้", output_stems_valid, "ใช้เป็นชื่อโฟลเดอร์และชื่อรายงาน"),
                            ("โหมดส่งออกพร้อม", export_mode_ready, ""),
                        ]
                    )
            with st.expander("ตัวอย่างไฟล์ใน Batch ZIP", expanded=False):
                st.code(
                    "\n".join(
                        [
                            "batch_summary.xlsx",
                            "<output_stem>/report.xlsx",
                            "<output_stem>/export_summary.txt",
                            "<output_stem>/session.tmcproj.json",
                            "<output_stem>/mapping.json",
                            "<output_stem>/charts/",
                        ]
                    ),
                    language="text",
                )
                st.caption("Batch ZIP ไม่รวม raw input Excel files และไม่รวม local file paths")
            if batch_result:
                _render_section_header("สถานะส่งออกรายไฟล์", "ผลการสร้างรายงานใน Batch ล่าสุด")
                status_display = _batch_status_display_frame(batch_analysis, batch_result)
                status_columns = [
                    column
                    for column in ["ชื่อไฟล์", "ชื่อส่งออก", "สถานะส่งออก", "โหมดส่งออกที่ใช้", "หมายเหตุ"]
                    if column in status_display.columns
                ]
                st.dataframe(status_display[status_columns] if status_columns else status_display, width="stretch", hide_index=True)
                with st.expander("ตัวอย่างไฟล์ใน Batch ZIP ล่าสุด", expanded=False):
                    st.code("\n".join(batch_zip_contents_preview(batch_result.summary_rows)), language="text")
                    st.caption("Batch ZIP ไม่รวม raw input Excel files และไม่รวม local file paths")
                st.download_button(
                    "ดาวน์โหลด Batch ZIP",
                    data=download_buffer(batch_result.package_bytes),
                    file_name=batch_package_filename(st.session_state.get("tmc_batch_preset_name") or batch_preset_name or "tmc_batch"),
                    mime=BATCH_PACKAGE_MIME,
                    key="download_batch_zip",
                )
            elif not batch_analysis:
                _render_empty_state("ยังส่งออก Batch ไม่ได้", "วิเคราะห์ Batch และกำหนด Peak ก่อนสร้าง ZIP")
            elif not peaks_ready:
                _render_alert(batch_zip_generation_block_reason(has_successful_files=True, peaks_ready=False, batch_stale=False), "warning")

        if active_tab == "ตรวจสอบข้อมูล":
            _render_section_header(
                "ตรวจสอบข้อมูล Batch",
                "ตรวจสอบสถานะรายไฟล์, QC รวม, และรายละเอียด Batch_QC ก่อนนำผลไปใช้ต่อ",
            )
            batch_analysis = st.session_state.get("tmc_batch_analysis_result")
            batch_result = st.session_state.get("tmc_batch_export_result")
            status_frame = _batch_status_frame(batch_analysis, batch_result)
            if batch_analysis and batch_stale:
                _render_alert("ข้อมูล Batch มีการเปลี่ยนแปลง กรุณาวิเคราะห์ Batch ใหม่", "warning")
            elif batch_analysis and batch_export_stale:
                _render_alert("ข้อมูลส่งออกมีการเปลี่ยนแปลง กรุณาสร้าง Batch ZIP ใหม่", "warning")

            if not batch_analysis:
                _render_empty_state("กรุณาวิเคราะห์ Batch ก่อนตรวจสอบผลรวม", "ผลรวมรายไฟล์และ Batch_QC จะแสดงหลังการวิเคราะห์ Batch")
            else:
                counts = _batch_summary_counts(status_frame)
                _render_metric_strip(
                    [
                        ("ไฟล์ทั้งหมด", f"{counts['total_files']:,}", "ไฟล์", ""),
                        ("สำเร็จ", f"{counts['successful_files']:,}", "ไฟล์", "", "success" if counts["successful_files"] else "รอตรวจ"),
                        ("ล้มเหลว", f"{counts['failed_files']:,}", "ไฟล์", "", "failed" if counts["failed_files"] else "พร้อม"),
                        ("QC ผิดพลาด", f"{counts['QC_errors']:,}", "รายการ", "", "error" if counts["QC_errors"] else "พร้อม"),
                        ("QC เตือน", f"{counts['QC_warnings']:,}", "รายการ", "", "warning" if counts["QC_warnings"] else "พร้อม"),
                        ("QC ข้อมูล", f"{counts['QC_info']:,}", "รายการ", "", "info" if counts["QC_info"] else "พร้อม"),
                    ],
                    columns=6,
                )

                _render_section_header("สถานะรายไฟล์", "ตารางตรวจสอบผลวิเคราะห์และความพร้อมส่งออกของแต่ละไฟล์")
                status_display = _batch_status_display_frame(batch_analysis, batch_result)
                status_columns = _existing_columns(
                    status_display,
                    [
                        "ชื่อไฟล์",
                        "วันที่สำรวจ",
                        "ชื่อส่งออก",
                        "สถานะ",
                        "AM กำหนดแล้ว",
                        "PM กำหนดแล้ว",
                        "PCU รวม",
                        "QC ผิดพลาด",
                        "QC เตือน",
                        "QC ข้อมูล",
                        "สถานะส่งออก",
                        "หมายเหตุ",
                    ],
                )
                st.dataframe(status_display[status_columns] if status_columns else status_display, width="stretch", hide_index=True)

                failed = status_frame[status_frame["status"].astype(str) == "failed"]
                if not failed.empty:
                    _render_alert("พบไฟล์ที่วิเคราะห์หรือส่งออกไม่สำเร็จ ไฟล์เหล่านี้ไม่ต้องกำหนด Peak", "warning")
                    failed_columns = _existing_columns(failed, ["file_name", "output_stem", "status", "notes"])
                    failed_display = failed[failed_columns] if failed_columns else failed
                    if "status" in failed_display.columns:
                        failed_display = failed_display.copy()
                        failed_display["status"] = failed_display["status"].map(_display_status_label)
                    st.dataframe(failed_display, width="stretch", hide_index=True)

                batch_qc = _batch_qc_rows_for_ui(batch_analysis, batch_result)
                if not batch_qc.empty:
                    _render_section_header("Batch_QC", "ตัวอย่างรายการ QC รวมสำหรับตรวจสอบก่อนใช้ผลต่อ")
                    batch_qc_display = _batch_qc_preview_display_frame(batch_qc)
                    st.dataframe(batch_qc_display.head(25), width="stretch", hide_index=True)
                    with st.expander("Batch_QC รายละเอียดทั้งหมด", expanded=False):
                        st.dataframe(_qc_display_frame(batch_qc, thai_labels=True), width="stretch", hide_index=True)
                else:
                    _render_alert("ยังไม่มีรายการ Batch_QC สำหรับไฟล์ที่วิเคราะห์สำเร็จ", "success")

                if batch_result:
                    _render_section_header("ร่องรอย Batch ZIP", "ตรวจสอบองค์ประกอบหลักของแพ็กเกจส่งออกล่าสุด")
                    _render_readiness_checklist(
                        [
                            ("สร้าง batch_summary.xlsx แล้ว", True, ""),
                            ("มี Sheet Batch_Summary", True, ""),
                            ("มี Sheet Batch_QC", True, ""),
                            ("ไม่รวม raw input Excel", True, ""),
                        ]
                    )

                with st.expander("สรุป QC รายไฟล์", expanded=False):
                    qc_summary_columns = _existing_columns(status_frame, ["file_name", "QC errors", "QC warnings", "QC info"])
                    qc_summary = status_frame[qc_summary_columns] if qc_summary_columns else status_frame
                    st.dataframe(_format_display_columns(qc_summary), width="stretch", hide_index=True)
                with st.expander("รายละเอียดวิเคราะห์ Batch", expanded=False):
                    if batch_analysis:
                        st.write(
                            [
                                {
                                    "file_name": item.file_name,
                                    "detected_sheets": item.detected_sheets,
                                    "matched_sheet_count": item.matched_sheet_count,
                                    "missing_detected_sheet_count": item.missing_detected_sheet_count,
                                    "extra_preset_row_count": item.extra_preset_row_count,
                                    "notes": item.notes,
                                }
                                for item in batch_analysis.items
                            ]
                        )

    processed = st.session_state.get("tmc_processed")
    pce_results_stale = _processed_pce_results_stale(processed, selected_pce_factors)
    if pce_results_stale:
        st.session_state["tmc_pce_results_stale"] = True
        st.session_state.pop("tmc_output", None)
    result = None if pce_results_stale else (processed["result"] if processed else None)
    mapping_df = processed["mapping"] if processed and not pce_results_stale else mapping
    hourly_movement = _hourly_movement_for_ui(result, mapping_df)
    confirmed_am_start = st.session_state.get("tmc_confirmed_am_peak_start", "")
    confirmed_am_end = st.session_state.get("tmc_confirmed_am_peak_end", "")
    confirmed_pm_start = st.session_state.get("tmc_confirmed_pm_peak_start", "")
    confirmed_pm_end = st.session_state.get("tmc_confirmed_pm_peak_end", "")

    if is_single_file_mode:
        if active_tab == "ตรวจ Peak":
            _render_section_header("ตรวจ Peak", "ตรวจสอบรูปแบบปริมาณจราจรรายชั่วโมง และกำหนดช่วง AM/PM Peak สำหรับใช้ในรายงาน")
            if pce_results_stale:
                _render_alert("ค่า PCE เปลี่ยนหลังจากประมวลผลแล้ว กรุณาประมวลผลใหม่ก่อนตรวจ Peak หรือส่งออกรายงาน", "warning")
            if result is None:
                _render_empty_state(
                    "กรุณาประมวลผลข้อมูลก่อนตรวจ Peak",
                    "เมื่อประมวลผลแล้ว ระบบจะแสดงกราฟ PCU รายชั่วโมงและตัวเลือกยืนยัน AM/PM Peak",
                )
            else:
                am_start, am_end, am_pcu = _peak_period_text(result.peaks, "AM")
                pm_start, pm_end, pm_pcu = _peak_period_text(result.peaks, "PM")
                _render_section_header("กราฟ PCU รายชั่วโมง", "ตรวจแนวโน้มปริมาณรวมก่อนเลือกช่วง Peak ที่ใช้เป็นค่าหลัก")
                _render_hourly_pcu_line_chart(
                    hourly_movement,
                    f"{am_start}-{am_end}" if am_start and am_end else "",
                    f"{pm_start}-{pm_end}" if pm_start and pm_end else "",
                )

                interval_options = _hourly_interval_options(hourly_movement, result.peaks)
                if interval_options:
                    option_labels = [label for label, _, _ in interval_options]
                    am_default = f"{am_start}-{am_end}" if am_start and am_end else option_labels[0]
                    pm_default = f"{pm_start}-{pm_end}" if pm_start and pm_end else option_labels[min(1, len(option_labels) - 1)]
                    loaded_confirmed_peaks = st.session_state.get("tmc_loaded_confirmed_peaks") or {}
                    loaded_am_label = f"{loaded_confirmed_peaks.get('am_peak_start', '')}-{loaded_confirmed_peaks.get('am_peak_end', '')}".strip("-")
                    loaded_pm_label = f"{loaded_confirmed_peaks.get('pm_peak_start', '')}-{loaded_confirmed_peaks.get('pm_peak_end', '')}".strip("-")
                    if loaded_am_label in option_labels and not st.session_state.get("am_peak_period_select"):
                        am_default = loaded_am_label
                    if loaded_pm_label in option_labels and not st.session_state.get("pm_peak_period_select"):
                        pm_default = loaded_pm_label
                    am_index = option_labels.index(am_default) if am_default in option_labels else 0
                    pm_index = option_labels.index(pm_default) if pm_default in option_labels else min(1, len(option_labels) - 1)

                    _render_section_header(
                        "กำหนดช่วง Peak",
                        "ระบบตรวจจับอัตโนมัติเป็นค่าแนะนำ ผู้ตรวจกำหนดช่วงที่จะใช้ในรายงาน",
                    )
                    confirm_cols = st.columns(2)
                    with confirm_cols[0]:
                        _render_peak_card("AM Peak · ระบบตรวจจับอัตโนมัติ", f"{am_start}-{am_end}" if am_start and am_end else "", am_pcu, "auto_suggested")
                        am_peak_label = st.selectbox("ช่วงที่กำหนด AM", option_labels, index=am_index, key="am_peak_period_select")
                    with confirm_cols[1]:
                        _render_peak_card("PM Peak · ระบบตรวจจับอัตโนมัติ", f"{pm_start}-{pm_end}" if pm_start and pm_end else "", pm_pcu, "auto_suggested")
                        pm_peak_label = st.selectbox("ช่วงที่กำหนด PM", option_labels, index=pm_index, key="pm_peak_period_select")
                    confirmed_am_start, confirmed_am_end = _selected_interval(interval_options, am_peak_label)
                    confirmed_pm_start, confirmed_pm_end = _selected_interval(interval_options, pm_peak_label)
                    previous_confirmed = (
                        st.session_state.get("tmc_confirmed_am_peak_start"),
                        st.session_state.get("tmc_confirmed_am_peak_end"),
                        st.session_state.get("tmc_confirmed_pm_peak_start"),
                        st.session_state.get("tmc_confirmed_pm_peak_end"),
                    )
                    current_confirmed = (confirmed_am_start, confirmed_am_end, confirmed_pm_start, confirmed_pm_end)
                    st.session_state["tmc_confirmed_am_peak_start"] = confirmed_am_start
                    st.session_state["tmc_confirmed_am_peak_end"] = confirmed_am_end
                    st.session_state["tmc_confirmed_pm_peak_start"] = confirmed_pm_start
                    st.session_state["tmc_confirmed_pm_peak_end"] = confirmed_pm_end
                    review_transition = _sync_single_workflow_from_state(
                        source_bytes=file_bytes if uploaded_file is not None else None,
                        source_file_name=uploaded_file.name if uploaded_file is not None else None,
                        export_mode=export_mode,
                    )
                    export_invalidated = review_transition.export_invalidated or (
                        previous_confirmed != current_confirmed and _clear_single_export()
                    )
                    if export_invalidated:
                        _flash_and_rerun("Peak เปลี่ยนแปลงแล้ว กรุณาสร้างรายงานใหม่")
                    confirmed_am_label = f"{confirmed_am_start}-{confirmed_am_end}" if confirmed_am_start and confirmed_am_end else ""
                    confirmed_pm_label = f"{confirmed_pm_start}-{confirmed_pm_end}" if confirmed_pm_start and confirmed_pm_end else ""
                    _render_metric_strip(
                        [
                            ("AM Peak", confirmed_am_label or "-", "", "ช่วงที่กำหนด", "กำหนดแล้ว" if confirmed_am_label else "รอตรวจสอบ"),
                            ("PM Peak", confirmed_pm_label or "-", "", "ช่วงที่กำหนด", "กำหนดแล้ว" if confirmed_pm_label else "รอตรวจสอบ"),
                            ("AM PCU", _interval_total_pcu(hourly_movement, confirmed_am_label) or am_pcu or "-", "PCU", "Peak PCU"),
                            ("PM PCU", _interval_total_pcu(hourly_movement, confirmed_pm_label) or pm_pcu or "-", "PCU", "Peak PCU"),
                        ],
                        columns=4,
                    )
                    _render_action_hint("ใช้ช่วงนี้เป็นค่าหลักสำหรับรายงาน")
                    if all([confirmed_am_start, confirmed_am_end, confirmed_pm_start, confirmed_pm_end]):
                        _render_alert("กำหนดช่วง Peak แล้ว พร้อมส่งออก", "success")
                    else:
                        _render_alert("กรุณากำหนด AM Peak และ PM Peak ก่อนส่งออก", "warning")
                else:
                    _render_alert("ไม่มีช่วงเวลารายชั่วโมงสำหรับกำหนด Peak", "warning")

                _render_section_header("สรุปทางเทคนิค", "แสดงเฉพาะค่าที่มีจากผลประมวลผลปัจจุบัน")
                _render_metric_strip(
                    [
                        ("จำนวนแถว", format_count(len(result.normalized)), "แถว", "normalized"),
                        ("จำนวนรถรวม", format_count(result.normalized["count"].sum()) if not result.normalized.empty else "0", "คัน", ""),
                        ("PCU รวม", format_pcu(result.normalized["pcu"].sum()) if not result.normalized.empty else "0", "PCU", ""),
                        ("QC", format_count(len(result.qc)), "", "ประเด็น"),
                    ],
                    columns=4,
                )
                _render_qc_status(result.qc)

                with st.expander("ตารางปริมาณจราจรแยกตามทิศทาง", expanded=False):
                    st.dataframe(_format_display_columns(hourly_movement), width="stretch", hide_index=True)
    
    if is_single_file_mode:
        if active_tab == "ส่งออก":
            export_peak_state = _single_effective_peak_state(result if result is not None and not pce_results_stale else None)
            effective_peaks = dict(export_peak_state.get("values") or {})
            confirmed_ready = bool(export_peak_state.get("ready"))
            _render_section_header("ส่งออกรายงาน", "สร้างรายงาน Excel และชุดไฟล์ประกอบสำหรับตรวจสอบย้อนหลัง")
            single_export_options = _single_export_mode_options(excel_com_status)
            previous_export_mode = st.session_state.get("report_export_mode", export_mode)
            selected_export_mode = st.radio(
                "โหมดส่งออกรายงาน",
                options=single_export_options,
                index=single_export_options.index(_coerce_export_mode(previous_export_mode, single_export_options, _default_single_export_mode(excel_com_status))),
                key="report_export_mode_control",
                horizontal=True,
                help="Excel Template Mode รักษา Native Chart และรูปแบบ Excel Template เมื่อ Excel COM พร้อมใช้งาน. Safe PNG Export Mode ใช้กราฟ PNG แบบคงที่.",
            )
            if apply_single_export_mode_change(selected_export_mode, previous_export_mode):
                st.rerun()
            export_mode = selected_export_mode
            use_template_report_layout = _use_template_layout_for_export(export_mode)
            use_excel_com_native_charts = _use_excel_native_charts_for_export(export_mode, excel_com_status)
            is_v2_single_result = _is_v2_result(result)
            if pce_results_stale:
                _render_alert("ผลลัพธ์เดิมไม่ตรงกับค่า PCE ปัจจุบัน ระบบปิดการส่งออกไว้จนกว่าจะประมวลผลใหม่", "warning")
            _render_action_hint("สร้างรายงานหลังจากประมวลผลและมีช่วงเร่งด่วน AM/PM พร้อมใช้งานแล้ว")
            _render_effective_peak_export_confirmation(result if result is not None else None, export_peak_state)
            export_run = st.button("สร้างรายงาน Excel", type="primary", disabled=not (result is not None and confirmed_ready and not pce_results_stale))

            export_status_col, readiness_col = st.columns([0.9, 1.1])
            with export_status_col:
                with st.container(border=True):
                    _render_section_header("โหมดส่งออก", "เลือกวิธีสร้างรายงานจากหน้านี้")
                    if export_mode == EXCEL_TEMPLATE_EXPORT_MODE:
                        st.markdown('<div class="tmc-mode-note tmc-mode-note-success"><strong>Excel Template Mode</strong> · แนะนำสำหรับรายงานฉบับใช้งานจริง</div>', unsafe_allow_html=True)
                        if is_v2_single_result:
                            st.caption(
                                "Excel Template Mode สำหรับ approach_movement ใช้ Excel COM เพื่อรักษากราฟและ diagram จาก template"
                                if use_excel_com_native_charts
                                else V2_EXCEL_TEMPLATE_MODE_BLOCK_MESSAGE
                            )
                        else:
                            st.caption("เหมาะสำหรับรายงานฉบับใช้งานจริง รักษา Native Chart และรูปแบบ Excel Template เมื่อ Excel COM พร้อมใช้งาน")
                    else:
                        st.markdown('<div class="tmc-mode-note tmc-mode-note-warning"><strong>Safe PNG Export Mode</strong> · โหมดสำรอง</div>', unsafe_allow_html=True)
                        st.caption("โหมดสำรอง เหมาะสำหรับตรวจร่างหรือกรณี Excel COM ใช้งานไม่ได้")

                    if excel_com_status.available:
                        version_text = f"Excel version: {excel_com_status.version}" if excel_com_status.version else "พร้อมใช้งาน"
                        _render_alert(f"Excel COM พร้อมใช้งาน: {version_text}", "success")
                    else:
                        _render_alert(f"Excel COM ไม่พร้อมใช้งาน ระบบจะใช้โหมดสำรองแบบ PNG: {excel_com_status.reason}", "warning")
                    st.caption("ใช้ปุ่มทดสอบ Excel COM ใน sidebar หากต้องการตรวจสถานะใหม่")
                    with st.expander("รายละเอียด Excel COM", expanded=False):
                        _render_excel_com_status(excel_com_status)

            with readiness_col:
                with st.container(border=True):
                    _render_section_header("ความพร้อมก่อนส่งออก", "รายการตรวจสอบแบบย่อก่อนสร้างรายงาน")
                    if result is not None:
                        _render_qc_status(result.qc)
                    _render_readiness_checklist(
                        [
                            ("โหลดไฟล์สำรวจแล้ว", uploaded_file is not None, ""),
                            ("Mapping พร้อมใช้งาน", bool(st.session_state.get("mapping_table")), ""),
                            ("ประมวลผลแล้ว", result is not None, "ค่า PCE เปลี่ยน กรุณาประมวลผลใหม่" if pce_results_stale else ""),
                            ("กำหนดช่วงเร่งด่วน AM/PM แล้ว", confirmed_ready, str(export_peak_state.get("summary_text") or "")),
                            (
                                "Excel COM พร้อมใช้งาน",
                                bool(excel_com_status.available) if export_mode == EXCEL_TEMPLATE_EXPORT_MODE else True,
                                V2_EXCEL_TEMPLATE_MODE_BLOCK_MESSAGE
                                if is_v2_single_result and export_mode == EXCEL_TEMPLATE_EXPORT_MODE and not excel_com_status.available
                                else ("จำเป็นสำหรับ Excel Template Mode" if export_mode == EXCEL_TEMPLATE_EXPORT_MODE else "ไม่จำเป็นในโหมดสำรอง"),
                            ),
                        ]
                    )

            if export_run:
                set_active_tab("ส่งออก")
                excel_com_requested = bool(use_excel_com_native_charts)
                export_excel_com_status = probe_excel_com() if excel_com_requested else None
                excel_com_enabled = bool(export_excel_com_status and export_excel_com_status.available)
                if excel_com_requested and export_excel_com_status is not None and not export_excel_com_status.available:
                    st.warning(
                        "Excel COM ไม่พร้อมใช้งาน ระบบจะใช้โหมดสำรองแบบ PNG "
                        f"สาเหตุ: {export_excel_com_status.reason}. {export_excel_com_status.detail}"
                    )
    
                confirmed_setup = {
                    **setup,
                    "am_peak_start": effective_peaks.get("am_peak_start", ""),
                    "am_peak_end": effective_peaks.get("am_peak_end", ""),
                    "pm_peak_start": effective_peaks.get("pm_peak_start", ""),
                    "pm_peak_end": effective_peaks.get("pm_peak_end", ""),
                    "peak_selection_source": str(export_peak_state.get("source") or PEAK_SELECTION_AUTO),
                }
                confirmed_periods = {
                    "AM": (str(effective_peaks.get("am_peak_start", "")), str(effective_peaks.get("am_peak_end", ""))),
                    "PM": (str(effective_peaks.get("pm_peak_start", "")), str(effective_peaks.get("pm_peak_end", ""))),
                }
                export_generated_at = generated_timestamp_text()
                try:
                    raw_sheets = {name: parsed.data for name, parsed in parsed_details.items()}
                    if is_v2_single_result:
                        confirmed_result = process_tmc_dry_run_v2(
                            raw_sheets=raw_sheets,
                            mapping=mapping_df,
                            setup={**confirmed_setup, "movement_code_scheme": MOVEMENT_SCHEME_V2},
                            detected_sheets=detected_sheet_names,
                            peak_mode=peak_mode,
                            peak_windows=peak_windows,
                            confirmed_peak_periods=confirmed_periods,
                            pce_factors=selected_pce_factors,
                        )
                        workbook_bytes = _export_single_file_for_ui(
                            result=confirmed_result,
                            mapping=mapping_df,
                            setup=confirmed_setup,
                            export_mode=export_mode,
                            use_template_report_layout=use_template_report_layout,
                            use_excel_com_native_charts=excel_com_enabled,
                            source_file_name=uploaded_file.name if uploaded_file is not None else st.session_state.get("tmc_loaded_source_file_name", ""),
                            generated_at=export_generated_at,
                        )
                        confirmed_result.workbook_bytes = workbook_bytes
                    else:
                        with warnings.catch_warnings(record=True) as export_warnings:
                            warnings.simplefilter("always", RuntimeWarning)
                            confirmed_result = process_tmc(
                                raw_sheets=raw_sheets,
                                mapping=mapping_df,
                                setup=confirmed_setup,
                                detected_sheets=detected_sheet_names,
                                peak_mode=peak_mode,
                                peak_windows=peak_windows,
                                confirmed_peak_periods=confirmed_periods,
                                pce_factors=selected_pce_factors,
                                generate_workbook=True,
                                use_template_report_layout=use_template_report_layout,
                                use_excel_com_native_charts=excel_com_enabled,
                                export_mode=export_mode,
                                source_file_name=uploaded_file.name if uploaded_file is not None else st.session_state.get("tmc_loaded_source_file_name", ""),
                                generated_at=export_generated_at,
                            )
                        for warning in export_warnings:
                            message = str(warning.message)
                            if excel_com_requested and "Excel COM native-chart export failed after COM was available" in message:
                                st.warning(f"Excel COM ส่งออก Native Chart ไม่สำเร็จ ระบบใช้โหมดสำรองแบบ PNG ({message})")
                            elif excel_com_requested and "Excel COM unavailable" in message:
                                st.warning(f"Excel COM ไม่พร้อมใช้งาน ระบบใช้โหมดสำรองแบบ PNG ({message})")
                    if not confirmed_result.workbook_bytes:
                        st.error("ส่งออกเสร็จแล้ว แต่ไฟล์ Excel ที่สร้างไม่มีข้อมูล")
                    else:
                        confirmed_hourly_movement = _hourly_movement_for_ui(confirmed_result, mapping_df)
                        chart_pngs = report_chart_pngs(
                            confirmed_hourly_movement,
                            vehicle_composition_report(confirmed_result.normalized),
                            setup=confirmed_setup,
                        )
                        diagram_png = b""
                        if not _is_v2_result(confirmed_result):
                            diagram_png = generate_four_leg_tmc_diagram(
                                confirmed_hourly_movement,
                                confirmed_result.peaks,
                                DiagramConfig(
                                    tmc_id=tmc_id,
                                    tmc_name=tmc_title,
                                    survey_date_text=survey_date_text,
                                    north_label=north_label,
                                    south_label=south_label,
                                    east_label=east_label,
                                    west_label=west_label,
                                    north_road=north_road,
                                    south_road=south_road,
                                    east_road=east_road,
                                    west_road=west_road,
                                    survey_period_text=survey_period,
                                    caption_text=caption_text,
                                    show_u_turn=show_u_turn,
                                ),
                            )
                        st.session_state["tmc_output"] = {
                            "result": confirmed_result,
                            "chart_pngs": dict(chart_pngs),
                            "diagram_png": diagram_png,
                            "workbook_bytes": confirmed_result.workbook_bytes,
                            "workbook_filename": safe_workbook_filename(tmc_id),
                            "confirmed_setup": confirmed_setup,
                            "export_mode": export_mode,
                            "generated_at": export_generated_at,
                        }
                        set_active_tab("ส่งออก")
                        _flash_and_rerun("สร้างรายงาน Excel เสร็จแล้ว")
                except Exception as exc:  # pragma: no cover - UI guardrail
                    st.error(f"ส่งออกไฟล์ไม่สำเร็จ: {exc}")
    
            output = st.session_state.get("tmc_output")
            if output:
                _render_section_header("ดาวน์โหลดและรายละเอียด", "ไฟล์ที่สร้างแล้วและชุดประกอบสำหรับตรวจสอบย้อนหลัง")
                _render_download_button("ดาวน์โหลดรายงาน Excel", output["workbook_bytes"], output["workbook_filename"], EXCEL_MIME)
                session_bytes = st.session_state.get("tmc_project_session_bytes")
                session_filename = st.session_state.get("tmc_project_session_filename", "tmc_session.tmcproj.json")
                if not session_bytes:
                    session = _build_session_from_state(
                        uploaded_file.name if uploaded_file is not None else None,
                        len(file_bytes) if uploaded_file is not None else None,
                    )
                    session_bytes = session_to_json_bytes(session)
                    filename_seed = st.session_state.get("tmc_id_input") or st.session_state.get("tmc_title_input") or (uploaded_file.name if uploaded_file is not None else None)
                    session_filename = safe_project_session_filename(filename_seed)
                    st.session_state["tmc_project_session_bytes"] = session_bytes
                    st.session_state["tmc_project_session_filename"] = session_filename

                output_result = output["result"]
                output_setup = output.get("confirmed_setup", setup)
                summary_text = build_export_summary_text(
                    setup=output_setup,
                    source_file_name=uploaded_file.name if uploaded_file is not None else st.session_state.get("tmc_loaded_source_file_name", ""),
                    export_mode=output.get("export_mode", export_mode),
                    peaks=output_result.peaks,
                    mapping=mapping_df,
                    qc=output_result.qc,
                    workbook_filename=output["workbook_filename"],
                    pce_factors=output_result.pce_factors,
                    export_settings={
                        "template_version": TEMPLATE_VERSION,
                        "template_name": Path(DEFAULT_TEMPLATE_PATH).name,
                        "template_map_name": Path(DEFAULT_TEMPLATE_MAP_PATH).name,
                    },
                    generated_at=output.get("generated_at"),
                )
                mapping_preset_bytes = serialize_mapping_preset(
                    build_mapping_preset(
                        mapping_df,
                        preset_name=str(st.session_state.get("tmc_id_input") or st.session_state.get("tmc_title_input") or "TMC Mapping Preset"),
                        movement_code_scheme=_current_mapping_scheme(),
                    )
                )
                if _is_v2_result(output_result):
                    package_bytes = create_v2_generated_export_package_zip(
                        workbook_bytes=output["workbook_bytes"],
                        workbook_filename=output["workbook_filename"],
                        setup=output_setup,
                        peaks=output_result.peaks,
                        mapping=mapping_df,
                        qc=output_result.qc,
                        mapping_preset_bytes=mapping_preset_bytes,
                        mapping_preset_filename="mapping_preset.mapping.json",
                        source_file_name=uploaded_file.name if uploaded_file is not None else st.session_state.get("tmc_loaded_source_file_name", ""),
                        export_mode=output.get("export_mode", export_mode),
                        generated_at=output.get("generated_at"),
                    )
                else:
                    package_bytes = create_export_package_zip(
                        workbook_bytes=output["workbook_bytes"],
                        workbook_filename=output["workbook_filename"],
                        export_summary_text=summary_text,
                        project_session_bytes=session_bytes,
                        project_session_filename=session_filename,
                        mapping_preset_bytes=mapping_preset_bytes,
                        mapping_preset_filename="mapping_preset.mapping.json",
                        mapping=mapping_df,
                        chart_pngs=output.get("chart_pngs", {}),
                        diagram_png=output.get("diagram_png"),
                    )
                with st.expander("ตัวอย่างไฟล์ใน Export Package ZIP", expanded=False):
                    if _is_v2_result(output_result):
                        preview_files = [
                            output["workbook_filename"],
                            "export_summary.txt",
                            "mapping_preset.mapping.json",
                            "mapping_table.xlsx",
                            "diagram/movement_diagram_data.csv",
                            "diagram/movement_diagram.png",
                        ]
                    else:
                        preview_files = [
                            output["workbook_filename"],
                            "export_summary.txt",
                            session_filename,
                            "mapping_preset.mapping.json",
                            "mapping_table.xlsx",
                            "charts/hourly_pcu_chart.png",
                            "charts/vehicle_composition_chart.png",
                            "charts/tmc_movement_diagram.png",
                        ]
                    st.code("\n".join(preview_files), language="text")
                    st.caption("Export Package ZIP ไม่รวม raw input Excel")
                _render_download_button(
                    "ดาวน์โหลด Export Package ZIP",
                    package_bytes,
                    safe_package_filename(output["workbook_filename"]),
                    PACKAGE_MIME,
                )
                if session_bytes:
                    st.download_button(
                        "ดาวน์โหลด Project Session",
                        data=download_buffer(session_bytes),
                        file_name=session_filename,
                        mime=PROJECT_SESSION_MIME,
                        key="download_project_session_export",
                    )
                with st.expander("กราฟและ Diagram สำหรับรายงาน", expanded=False):
                    chart_pngs = output["chart_pngs"]
                    chart_cols = st.columns(2)
                    with chart_cols[0]:
                        st.image(chart_pngs["hourly_pcu"], caption="ปริมาณจราจรรวมรายชั่วโมง")
                        _render_download_button("ดาวน์โหลดกราฟ PCU รายชั่วโมง (PNG)", chart_pngs["hourly_pcu"], "hourly_pcu_chart.png", PNG_MIME)
                    with chart_cols[1]:
                        st.image(chart_pngs["vehicle_composition"], caption="สัดส่วนประเภทยานพาหนะ")
                        _render_download_button(
                            "ดาวน์โหลดกราฟสัดส่วนยานพาหนะ (PNG)",
                            chart_pngs["vehicle_composition"],
                            "vehicle_composition_chart.png",
                            PNG_MIME,
                        )
                    if output.get("diagram_png"):
                        st.image(output["diagram_png"], caption="Four-leg TMC movement diagram")
                        _render_download_button("ดาวน์โหลด Diagram movement (PNG)", output["diagram_png"], "tmc_movement_diagram.png", PNG_MIME)
                    elif _is_v2_result(output_result):
                        st.caption("approach_movement diagram PNG อยู่ใน Export Package ZIP ที่ path diagram/movement_diagram.png")
            else:
                _render_empty_state(
                    "ยังไม่มีไฟล์ส่งออก",
                    "กำหนดช่วงเร่งด่วน AM/PM แล้วสร้างรายงาน Excel เมื่อพร้อม",
                )

    if is_single_file_mode:
        if active_tab == "ตรวจสอบข้อมูล":
            _render_section_header(
                "ตรวจสอบข้อมูล",
                "ตรวจสอบ QC, ข้อมูลที่ประมวลผลแล้ว และรายการ Audit ก่อนนำผลไปใช้ในรายงาน",
            )
            if result is None:
                _render_empty_state(
                    "กรุณาประมวลผลข้อมูลก่อนตรวจสอบรายละเอียด",
                    "QC, ตารางข้อมูล และรายการ Audit จะแสดงหลังมีผลประมวลผลแล้ว",
                )
            else:
                qc_counts = _qc_severity_counts(result.qc)
                total_qc = sum(qc_counts.values())
                _render_metric_strip(
                    [
                        ("QC ผิดพลาด", f"{qc_counts['error']:,}", "รายการ", "", "error" if qc_counts["error"] else "พร้อม"),
                        ("QC เตือน", f"{qc_counts['warning']:,}", "รายการ", "", "warning" if qc_counts["warning"] else "พร้อม"),
                        ("QC ข้อมูล", f"{qc_counts['info']:,}", "รายการ", "", "info" if qc_counts["info"] else "พร้อม"),
                        ("QC รวม", f"{total_qc:,}", "รายการ", "", "พร้อม" if not total_qc else "ต้องตรวจ"),
                    ],
                    columns=4,
                )
                if qc_counts["error"]:
                    _render_alert("พบ QC error กรุณาตรวจรายละเอียดก่อนส่งออก", "error")
                elif qc_counts["warning"]:
                    _render_alert("พบ QC warning ยังส่งออกได้ แต่ควรตรวจหมายเหตุก่อน", "warning")
                elif qc_counts["info"]:
                    _render_alert("มีหมายเหตุ QC info สำหรับตรวจสอบ", "info")
                else:
                    _render_alert("ไม่พบประเด็น QC", "success")

                _render_section_header("รายละเอียด QC", "รายการตรวจสอบที่ใช้ประกอบการพิจารณาก่อนส่งออกรายงาน")
                qc_display = _qc_display_frame(result.qc, thai_labels=True)
                if qc_display.empty:
                    _render_alert("ไม่พบรายการ QC ที่ต้องตรวจสอบ", "success")
                elif len(qc_display) > 25:
                    st.dataframe(qc_display.head(25), width="stretch", hide_index=True)
                    with st.expander("รายการ QC ทั้งหมด", expanded=False):
                        st.dataframe(qc_display, width="stretch", hide_index=True)
                else:
                    st.dataframe(qc_display, width="stretch", hide_index=True)

                with st.expander("Normalized Data", expanded=False):
                    st.dataframe(_format_display_columns(result.normalized.head(1000)), width="stretch", hide_index=True)
                with st.expander("Hourly Movement PCU", expanded=False):
                    if not hourly_movement.empty:
                        st.dataframe(_format_display_columns(hourly_movement), width="stretch", hide_index=True)
                    if not result.hourly.empty:
                        st.caption("Hourly raw summary")
                        st.dataframe(_format_display_columns(result.hourly), width="stretch", hide_index=True)
                with st.expander("Movement Summary", expanded=False):
                    st.dataframe(_format_display_columns(result.movement), width="stretch", hide_index=True)
                if _is_v2_result(result):
                    with st.expander("Movement Diagram Data", expanded=False):
                        diagram_data = build_v2_movement_diagram_data(
                            movement_summary=result.movement,
                            hourly_movement_pcu=hourly_movement,
                            peaks=result.peaks,
                        )
                        st.dataframe(_format_display_columns(diagram_data), width="stretch", hide_index=True)
                    with st.expander("Movement Code Reference", expanded=False):
                        st.dataframe(_v2_movement_code_reference_frame(), width="stretch", hide_index=True)
                with st.expander("Peak / PHF Data", expanded=False):
                    peak_display = _format_display_columns(result.peaks)
                    if "peak_start" in peak_display.columns:
                        peak_display["peak_start"] = peak_display["peak_start"].map(lambda value: str(value)[:5] if str(value) else "-")
                    if "peak_end" in peak_display.columns:
                        peak_display["peak_end"] = peak_display["peak_end"].map(lambda value: str(value)[:5] if str(value) else "-")
                    st.dataframe(peak_display, width="stretch", hide_index=True)
                with st.expander("Movement Aggregation Audit", expanded=False):
                    st.caption("ตารางตรวจสอบ source movement, source stream และ output movement ที่ใช้รวมค่าในรายงาน")
                    audit_frame = movement_aggregation_audit(result.normalized, mapping_df)
                    st.dataframe(_format_display_columns(audit_frame), width="stretch", hide_index=True)
                with st.expander("รายละเอียดการอ่านไฟล์และ Parser", expanded=False):
                    if uploaded_file is not None:
                        st.dataframe(preview_summary, width="stretch")
                        for sheet_name, preview in previews.items():
                            st.markdown(f"**{sheet_name}**")
                            st.dataframe(preview, width="stretch")
                        for sheet_name, parsed in parsed_details.items():
                            debug = parsed.debug
                            st.markdown(f"**Parser: {sheet_name}**")
                            st.write(
                                {
                                    "detected_first_data_row": debug.first_data_row,
                                    "detected_time_columns": {
                                        "time_start_col": debug.time_start_col,
                                        "time_end_col": debug.time_end_col,
                                    },
                                    "detected_vehicle_class_columns": debug.vehicle_class_columns,
                                }
                            )
                            st.dataframe(parsed.data.head(10), width="stretch")
                    else:
                        _render_action_hint("ไม่มี Workbook ที่อัปโหลดในรอบการทำงานนี้")
                with st.expander("Export Metadata / Template Diagnostics", expanded=False):
                    output = st.session_state.get("tmc_output")
                    if output:
                        st.write(
                            {
                                "workbook_filename": output.get("workbook_filename", ""),
                                "export_mode": output.get("export_mode", ""),
                                "generated_at": output.get("generated_at"),
                                "template_version": TEMPLATE_VERSION,
                                "template_name": Path(DEFAULT_TEMPLATE_PATH).name,
                                "template_map_name": Path(DEFAULT_TEMPLATE_MAP_PATH).name,
                            }
                        )
                    else:
                        _render_action_hint("ยังไม่มี metadata การส่งออกในรอบนี้")
                    _render_template_audit_notes()
                    _render_excel_com_status(excel_com_status)
    with top_status_placeholder.container():
        _render_top_status_bar(
            is_single_file_mode=is_single_file_mode,
            uploaded_name=uploaded_file.name if uploaded_file is not None else None,
            uploaded_count=len(batch_uploads or []),
            batch_mapping_ready=loaded_batch_preset is not None,
            export_mode=export_mode if is_single_file_mode else shell_batch_export_mode,
            excel_com_status=excel_com_status,
        )
    with workflow_shell_placeholder.container():
        _render_workflow_shell(
            is_single_file_mode=is_single_file_mode,
            uploaded_name=uploaded_file.name if uploaded_file is not None else None,
            uploaded_count=len(batch_uploads or []),
            batch_mapping_ready=loaded_batch_preset is not None,
            batch_signature=shell_batch_signature,
            export_mode=export_mode,
            excel_com_status=excel_com_status,
        )


if __name__ == "__main__":
    _run_streamlit_app()
    st.stop()
