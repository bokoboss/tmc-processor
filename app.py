"""Streamlit UI for the TMC Processor MVP."""

from __future__ import annotations

from datetime import time
import hashlib
from html import escape
from io import BytesIO
from pathlib import Path
import sys
import warnings

import pandas as pd
import streamlit as st

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
from tmc_processor.diagram import DiagramConfig, generate_four_leg_tmc_diagram
from tmc_processor.downloads import EXCEL_MIME, PNG_MIME, download_buffer, safe_workbook_filename
from tmc_processor.excel_com_export import ExcelComStatus, probe_excel_com
from tmc_processor.export_package import (
    PACKAGE_MIME,
    build_export_summary_text,
    create_export_package_zip,
    safe_package_filename,
)
from tmc_processor.importer import detect_raw_direction_sheet_names, load_detected_sheet_details, preview_detected_sheets
from tmc_processor.mapping import (
    apply_saved_mapping_to_sheets,
    default_mapping_for_sheets,
    mapping_to_excel_bytes,
    mapping_control_warnings,
    movement_aggregation_messages,
    read_mapping_excel,
    selectbox_options_with_existing_values,
    validate_mapping_for_processing,
)
from tmc_processor.mapping_preset import (
    MAPPING_PRESET_MIME,
    MappingPresetError,
    apply_mapping_preset_to_detected_sheets,
    build_mapping_preset,
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
from tmc_processor import metadata as setup_metadata
from tmc_processor.pcu import (
    get_default_pce_factors,
    normalize_pce_factors,
    pce_factor_traceability_frame,
    pce_factors_equal,
    validate_pce_factors,
)
from tmc_processor.peaks import PEAK_SELECTION_USER_CONFIRMED
from tmc_processor.pipeline import process_tmc
from tmc_processor.report_template import DEFAULT_TEMPLATE_MAP_PATH, DEFAULT_TEMPLATE_PATH, load_template_map
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


DEFAULT_RESPONSIBLE_PARTY = "ที่ปรึกษา"
DEFAULT_SURVEY_PERIOD = "7.00 - 19.00"
DEFAULT_WEATHER = "แจ่มใส"
DEFAULT_CAPTION_TEXT = "ปริมาณจราจรบนทางแยก PCU/12 ชม. (07.00-19.00 น.)"


DEFAULT_RESPONSIBLE_PARTY = setup_metadata.DEFAULT_RESPONSIBLE_PARTY
DEFAULT_SURVEY_PERIOD = setup_metadata.DEFAULT_SURVEY_PERIOD
DEFAULT_WEATHER = setup_metadata.DEFAULT_WEATHER
DEFAULT_CAPTION_TEXT = setup_metadata.DEFAULT_CAPTION_TEXT

EXCEL_TEMPLATE_EXPORT_MODE = "Excel Template Mode — แนะนำ"
SAFE_PNG_EXPORT_MODE = "Safe PNG Export Mode — โหมดสำรอง"


def _default_text_from_filename(filename: str | None) -> str:
    if not filename:
        return ""
    return Path(filename).stem.replace("_", " ").replace("-", " ").strip()


def _time_from_text(value: str) -> time:
    hour, minute = value.split(":")
    return time(int(hour), int(minute))


def _time_text(value: time) -> str:
    return value.strftime("%H:%M")


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
    return start_text, end_text, f"{pcu:,.0f}"


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
            --tmc-primary: #1f3a5f;
            --tmc-primary-dark: #0f2340;
            --tmc-primary-soft: #e8f0f8;
            --tmc-bg: #f8fafc;
            --tmc-surface: #ffffff;
            --tmc-surface-muted: #f1f5f9;
            --tmc-border: #e2e8f0;
            --tmc-border-strong: #cbd5e1;
            --tmc-text: #0f172a;
            --tmc-muted: #64748b;
            --tmc-success: #166534;
            --tmc-warning: #92400e;
            --tmc-error: #991b1b;
            --tmc-slate: #475569;

            --tmc-font-family: "Prompt", "Sarabun", "Segoe UI", Tahoma, sans-serif;
            --tmc-font-xs: 0.75rem;
            --tmc-font-sm: 0.86rem;
            --tmc-font-md: 0.95rem;
            --tmc-font-lg: 1.08rem;
            --tmc-font-xl: 1.9rem;

            --tmc-space-1: 0.25rem;
            --tmc-space-2: 0.5rem;
            --tmc-space-3: 0.75rem;
            --tmc-space-4: 1rem;
            --tmc-space-5: 1.25rem;
            --tmc-space-6: 1.5rem;

            --tmc-radius-sm: 6px;
            --tmc-radius-md: 8px;
            --tmc-radius-lg: 10px;
            --tmc-shadow-subtle: 0 1px 2px rgba(15, 23, 42, 0.04);
            --tmc-shadow-card: 0 8px 22px rgba(15, 23, 42, 0.06);
        }
        html, body, [class*="css"] {
            font-family: var(--tmc-font-family);
        }
        .stApp {
            background: var(--tmc-bg);
            color: var(--tmc-text);
            font-family: var(--tmc-font-family);
        }
        .block-container {
            max-width: 1320px;
            padding-top: var(--tmc-space-5);
            padding-bottom: 2.5rem;
        }
        .tmc-header {
            padding: var(--tmc-space-1) 0 var(--tmc-space-4) 0;
        }
        .tmc-title {
            color: var(--tmc-text);
            font-size: var(--tmc-font-xl);
            font-weight: 720;
            letter-spacing: 0;
            margin: 0;
        }
        .tmc-subtitle {
            color: var(--tmc-slate);
            font-size: var(--tmc-font-md);
            margin-top: var(--tmc-space-1);
        }
        .tmc-workflow {
            color: var(--tmc-muted);
            font-size: var(--tmc-font-sm);
            margin-top: var(--tmc-space-2);
        }
        .tmc-status-grid {
            display: grid;
            grid-template-columns: repeat(3, minmax(0, 1fr));
            gap: var(--tmc-space-3);
            margin: var(--tmc-space-2) 0 var(--tmc-space-5) 0;
        }
        .tmc-card {
            background: var(--tmc-surface);
            border: 1px solid var(--tmc-border);
            border-radius: var(--tmc-radius-md);
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
            color: var(--tmc-muted);
            font-size: var(--tmc-font-xs);
            font-weight: 650;
            letter-spacing: 0;
            line-height: 1.25;
        }
        .tmc-card-value {
            color: var(--tmc-text);
            font-size: var(--tmc-font-md);
            font-weight: 700;
            line-height: 1.28;
            overflow-wrap: anywhere;
        }
        .tmc-card-note {
            color: var(--tmc-muted);
            font-size: var(--tmc-font-xs);
            line-height: 1.25;
            overflow-wrap: anywhere;
        }
        .tmc-status-card .tmc-card-label,
        .tmc-status-card .tmc-card-value,
        .tmc-status-card .tmc-card-note {
            max-width: 100%;
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
        }
        .tmc-status-gray { border-left: 3px solid var(--tmc-border-strong); }
        .tmc-status-blue { border-left: 3px solid var(--tmc-primary); }
        .tmc-status-green { border-left: 3px solid var(--tmc-success); }
        .tmc-status-amber { border-left: 3px solid var(--tmc-warning); }
        .tmc-status-red { border-left: 3px solid var(--tmc-error); }

        .tmc-section-header {
            display: flex;
            gap: var(--tmc-space-3);
            align-items: flex-start;
            margin: var(--tmc-space-1) 0 var(--tmc-space-4) 0;
        }
        .tmc-section-accent {
            width: 4px;
            min-height: 2.1rem;
            border-radius: 999px;
            background: var(--tmc-primary);
            margin-top: 0.08rem;
        }
        .tmc-section-title {
            color: var(--tmc-text);
            font-size: var(--tmc-font-lg);
            font-weight: 720;
            line-height: 1.25;
            margin: 0;
        }
        .tmc-section-description {
            color: var(--tmc-muted);
            font-size: var(--tmc-font-sm);
            line-height: 1.35;
            margin-top: 0.18rem;
        }
        .tmc-section {
            background: var(--tmc-surface);
            border: 1px solid var(--tmc-border);
            border-radius: var(--tmc-radius-md);
            padding: var(--tmc-space-4);
            margin: var(--tmc-space-3) 0 var(--tmc-space-4) 0;
        }
        .tmc-empty-state {
            background: var(--tmc-surface);
            border: 1px dashed var(--tmc-border-strong);
            border-radius: var(--tmc-radius-md);
            padding: var(--tmc-space-5);
            color: var(--tmc-muted);
            margin: var(--tmc-space-3) 0 var(--tmc-space-4) 0;
        }
        .tmc-empty-title {
            color: var(--tmc-text);
            font-size: var(--tmc-font-md);
            font-weight: 700;
            margin-bottom: var(--tmc-space-1);
        }
        .tmc-empty-description {
            color: var(--tmc-muted);
            font-size: var(--tmc-font-sm);
            line-height: 1.45;
        }
        .tmc-peak-card {
            min-height: 116px;
            border-top: 3px solid var(--tmc-slate);
        }
        .tmc-peak-suggested {
            border-top-color: var(--tmc-slate);
            background: linear-gradient(180deg, #ffffff 0%, #f8fafc 100%);
        }
        .tmc-peak-confirmed {
            border-top-color: var(--tmc-primary);
            background: linear-gradient(180deg, #ffffff 0%, var(--tmc-primary-soft) 100%);
        }
        .tmc-peak-topline {
            display: flex;
            justify-content: space-between;
            gap: var(--tmc-space-2);
            align-items: flex-start;
            margin-bottom: var(--tmc-space-2);
        }
        .tmc-peak-badge {
            border-radius: 999px;
            padding: 0.1rem 0.48rem;
            font-size: 0.68rem;
            font-weight: 700;
            line-height: 1.4;
            white-space: nowrap;
        }
        .tmc-peak-suggested .tmc-peak-badge {
            color: var(--tmc-slate);
            background: var(--tmc-surface-muted);
        }
        .tmc-peak-confirmed .tmc-peak-badge {
            color: var(--tmc-primary-dark);
            background: var(--tmc-primary-soft);
            border: 1px solid #d2e3f3;
        }
        .tmc-peak-time {
            color: var(--tmc-text);
            font-size: 1.32rem;
            font-weight: 760;
            line-height: 1.25;
            overflow-wrap: anywhere;
        }
        .tmc-peak-pcu {
            color: var(--tmc-muted);
            font-size: var(--tmc-font-sm);
            margin-top: var(--tmc-space-1);
        }
        div[data-testid="stMetric"] {
            background: var(--tmc-surface);
            border: 1px solid var(--tmc-border);
            border-radius: var(--tmc-radius-md);
            padding: var(--tmc-space-3) var(--tmc-space-4);
            box-shadow: var(--tmc-shadow-subtle);
        }
        section[data-testid="stSidebar"] {
            background: #f3f6f9;
            border-right: 1px solid var(--tmc-border);
        }
        section[data-testid="stSidebar"] h2,
        section[data-testid="stSidebar"] h3 {
            color: var(--tmc-primary);
            letter-spacing: 0;
        }
        section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p {
            color: var(--tmc-muted);
            line-height: 1.45;
        }
        div[data-testid="stVerticalBlockBorderWrapper"] {
            border-color: var(--tmc-border);
            border-radius: var(--tmc-radius-lg);
            box-shadow: var(--tmc-shadow-subtle);
        }
        div[data-testid="stTabs"] [data-baseweb="tab-list"] {
            gap: var(--tmc-space-1);
            border-bottom: 1px solid var(--tmc-border);
        }
        div[data-testid="stTabs"] [data-baseweb="tab"] {
            color: var(--tmc-muted);
            border-radius: var(--tmc-radius-md) var(--tmc-radius-md) 0 0;
            padding: 0.42rem 0.95rem;
            margin-bottom: -1px;
            border-bottom: 3px solid transparent;
            font-weight: 650;
        }
        div[data-testid="stTabs"] [data-baseweb="tab"][aria-selected="true"] {
            color: var(--tmc-primary-dark);
            background: var(--tmc-primary-soft);
            border-bottom-color: var(--tmc-primary);
            font-weight: 760;
        }
        div[data-testid="stTabs"] [data-baseweb="tab"]:hover {
            color: var(--tmc-primary);
            background: var(--tmc-surface-muted);
        }
        div[data-testid="stTabs"] [data-baseweb="tab-highlight"] {
            background-color: transparent;
        }
        h2, h3, h4 {
            color: var(--tmc-text);
            letter-spacing: 0;
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
            color: #334155;
            font-weight: 650;
        }
        div.stButton > button,
        div.stDownloadButton > button {
            border-radius: var(--tmc-radius-sm);
            border-color: var(--tmc-border-strong);
            color: var(--tmc-text);
            min-height: 2.35rem;
            box-shadow: none;
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
            background: var(--tmc-primary-dark);
            border-color: var(--tmc-primary-dark);
            color: #ffffff;
        }
        div.stButton > button:disabled,
        div.stDownloadButton > button:disabled,
        button[data-testid="stBaseButton-primary"]:disabled {
            background: var(--tmc-border) !important;
            border-color: var(--tmc-border) !important;
            color: #94a3b8 !important;
            box-shadow: none;
        }
        .tmc-checklist {
            background: var(--tmc-surface);
            border: 1px solid var(--tmc-border);
            border-radius: var(--tmc-radius-md);
            box-shadow: var(--tmc-shadow-subtle);
            margin: var(--tmc-space-2) 0 var(--tmc-space-4) 0;
            overflow: hidden;
        }
        .tmc-check-item {
            display: flex;
            align-items: center;
            gap: var(--tmc-space-2);
            background: transparent;
            border-bottom: 1px solid var(--tmc-border);
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
            border-radius: 999px;
            flex: 0 0 auto;
            font-weight: 800;
            font-size: 0.72rem;
        }
        .tmc-check-ready .tmc-check-icon {
            color: var(--tmc-success);
            background: #ecfdf5;
        }
        .tmc-check-warn .tmc-check-icon {
            color: var(--tmc-warning);
            background: #fffbeb;
        }
        .tmc-check-body {
            flex: 1 1 auto;
            min-width: 0;
        }
        .tmc-check-label {
            font-weight: 650;
            line-height: 1.3;
        }
        .tmc-check-status {
            color: var(--tmc-muted);
            font-size: var(--tmc-font-xs);
            line-height: 1.3;
        }
        .tmc-check-detail {
            color: var(--tmc-muted);
            font-size: var(--tmc-font-xs);
            line-height: 1.35;
        }
        .tmc-mode-note {
            color: var(--tmc-muted);
            font-size: var(--tmc-font-sm);
            border: 1px solid var(--tmc-border);
            border-radius: var(--tmc-radius-md);
            padding: var(--tmc-space-2) var(--tmc-space-3);
            margin-bottom: var(--tmc-space-3);
        }
        .tmc-mode-note-success {
            color: var(--tmc-primary-dark);
            background: var(--tmc-primary-soft);
            border-color: #d2e3f3;
            border-left: 3px solid var(--tmc-primary);
        }
        .tmc-mode-note-warning {
            color: var(--tmc-warning);
            background: #fffbeb;
            border-color: #fde68a;
            border-left: 3px solid var(--tmc-warning);
        }
        .tmc-sidebar-badge {
            background: var(--tmc-surface);
            border: 1px solid var(--tmc-border);
            border-radius: var(--tmc-radius-md);
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
            color: var(--tmc-muted);
            font-size: var(--tmc-font-xs);
            line-height: 1.35;
            margin-top: var(--tmc-space-3);
        }
        div[data-testid="stAlert"] {
            border-radius: var(--tmc-radius-md);
        }
        @media (max-width: 900px) {
            .tmc-status-grid {
                grid-template-columns: repeat(2, minmax(0, 1fr));
            }
        }
        @media (max-width: 640px) {
            .tmc-status-grid {
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


def _status_tone(status: str) -> str:
    text = str(status).casefold()
    if text in {"complete", "confirmed", "ready", "loaded", "เสร็จสิ้น", "ยืนยันแล้ว", "พร้อมใช้งาน", "โหลดแล้ว"}:
        return "green"
    if text in {"needs review", "ต้องตรวจสอบ", "active", "กำลังใช้งาน"}:
        return "blue"
    return "gray"


def _status_card_html(label: str, status: str, note: str = "", tone: str | None = None) -> str:
    tone = tone or _status_tone(status)
    note_html = f'<div class="tmc-card-note">{escape(note)}</div>' if note else '<div class="tmc-card-note">&nbsp;</div>'
    return (
        f'<div class="tmc-card tmc-status-card tmc-status-{tone}">'
        f'<div class="tmc-card-label">{escape(label)}</div>'
        f'<div class="tmc-card-value">{escape(status)}</div>'
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
    peak_status = "ยืนยันแล้ว" if all(
        confirmed.get(key) for key in ("am_peak_start", "am_peak_end", "pm_peak_start", "pm_peak_end")
    ) else ("ต้องตรวจสอบ" if processed else "ยังไม่ได้โหลด")
    export_status = "พร้อมใช้งาน" if output_ready or (processed and peak_status == "ยืนยันแล้ว" and excel_ready) else "ยังไม่ได้โหลด"

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


def _probe_excel_com_for_ui(force: bool = False) -> ExcelComStatus:
    status_key = "tmc_excel_com_status"
    if force or status_key not in st.session_state:
        st.session_state[status_key] = probe_excel_com()
    return st.session_state[status_key]


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
    version = int(st.session_state.get("pce_editor_version", 0) or 0)
    editor_key = f"pce_factors_editor_{version}"
    return st.session_state.get(editor_key, st.session_state.get("pce_factors_table"))


def _current_pce_factors_from_state() -> dict[str, float]:
    _ensure_pce_factor_state()
    return normalize_pce_factors(_current_pce_factor_input())


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
        records = _pce_factor_records(selected)
        if st.session_state.get("pce_factors_table") != records:
            if "tmc_selected_pce_factors" in st.session_state and st.session_state["tmc_selected_pce_factors"] != selected:
                _mark_processed_outputs_stale_for_pce_change()
            st.session_state["pce_factors_table"] = records
        st.session_state["tmc_selected_pce_factors"] = selected

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


def _build_session_from_state(uploaded_name: str | None, uploaded_size: int | None) -> dict[str, object]:
    peak_settings = {
        "peak_mode": _state_value("peak_mode_select", DEFAULT_PEAK_MODE),
        "am_peak_window_start": _state_value("am_peak_window_start_input", _time_from_text(AM_WINDOW[0])),
        "am_peak_window_end": _state_value("am_peak_window_end_input", _time_from_text(AM_WINDOW[1])),
        "pm_peak_window_start": _state_value("pm_peak_window_start_input", _time_from_text(PM_WINDOW[0])),
        "pm_peak_window_end": _state_value("pm_peak_window_end_input", _time_from_text(PM_WINDOW[1])),
        **_confirmed_peaks_from_state(),
    }
    return build_project_session(
        metadata={
            "project_name": _state_value("project_name_input"),
            "tmc_id": _state_value("tmc_id_input"),
            "tmc_title": _state_value("tmc_title_input"),
            "survey_point": _state_value("survey_point_input"),
            "survey_date_text": _state_value("survey_date_text_input"),
            "weather": _state_value("weather_input"),
            "responsible_party": _state_value("responsible_party_input"),
            "survey_period": _state_value("survey_period_input"),
        },
        directions={
            "north_label": _state_value("north_label_input"),
            "south_label": _state_value("south_label_input"),
            "east_label": _state_value("east_label_input"),
            "west_label": _state_value("west_label_input"),
            "north_road": _state_value("north_road_input"),
            "south_road": _state_value("south_road_input"),
            "east_road": _state_value("east_road_input"),
            "west_road": _state_value("west_road_input"),
            "caption_text": _state_value("caption_text_input"),
        },
        mapping=_current_mapping_for_session(),
        detected_sheet_names=st.session_state.get("tmc_detected_sheet_names", []),
        pce_factors=_current_pce_factors_from_state(),
        peak_settings=peak_settings,
        export_settings={
            "use_template_report_layout": bool(_state_value("use_template_report_layout_checkbox", True)),
            "use_excel_com_native_charts": bool(_state_value("use_excel_com_native_charts_checkbox", False)),
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
            for stale_key in ["tmc_processed", "tmc_output", "tmc_pce_results_stale", "am_peak_period_select", "pm_peak_period_select"]:
                st.session_state.pop(stale_key, None)
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


def _interval_total_pcu(hourly_movement: pd.DataFrame, label: str) -> str:
    if hourly_movement.empty or not label or "Total" not in hourly_movement.columns:
        return ""
    time_column = str(hourly_movement.columns[0])
    rows = hourly_interval_rows(hourly_movement)
    matched = rows[rows[time_column].astype(str).str.replace(".", ":", regex=False) == label.replace(".", ":")]
    if matched.empty:
        return ""
    value = pd.to_numeric(matched.iloc[0].get("Total"), errors="coerce")
    return f"{value:,.0f}" if pd.notna(value) else ""


def _render_peak_card(title: str, period_label: str, pcu: str, source: str) -> None:
    is_confirmed = source == "user_confirmed"
    card_class = "tmc-peak-confirmed" if is_confirmed else "tmc-peak-suggested"
    badge = "ยืนยันแล้ว" if is_confirmed else "แนะนำ"
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

    with st.sidebar:
        st.header("ข้อมูลนำเข้า")
        uploaded_file = st.file_uploader(
            "อัปโหลดไฟล์ TMC Excel",
            type=["xlsx", "xlsm", "xls"],
            key="raw_tmc_upload",
        )

    file_bytes = uploaded_file.getvalue() if uploaded_file is not None else b""
    uploaded_identity = (uploaded_file.name, len(file_bytes)) if uploaded_file is not None else None
    uploaded_filename_default = _default_text_from_filename(uploaded_file.name if uploaded_file is not None else None)

    if uploaded_identity and st.session_state.get("tmc_uploaded_identity") != uploaded_identity:
        st.session_state.pop("tmc_output", None)
        st.session_state.pop("tmc_processed", None)
        st.session_state.pop("tmc_pce_results_stale", None)
        if not st.session_state.get("tmc_mapping_table_from_session"):
            st.session_state.pop("mapping_table", None)
        for key in list(st.session_state.keys()):
            if str(key).startswith("mapping_editor_"):
                st.session_state.pop(key, None)
        for key in ["am_peak_period_select", "pm_peak_period_select"]:
            st.session_state.pop(key, None)
        st.session_state["tmc_uploaded_identity"] = uploaded_identity
        if uploaded_filename_default:
            if not str(st.session_state.get("tmc_title_input", "")).strip():
                st.session_state["tmc_title_input"] = uploaded_filename_default
            if not str(st.session_state.get("survey_point_input", "")).strip():
                st.session_state["survey_point_input"] = uploaded_filename_default

    st.session_state.setdefault("project_name_input", "")
    st.session_state.setdefault("tmc_id_input", "")
    st.session_state.setdefault("tmc_title_input", uploaded_filename_default)
    st.session_state.setdefault("survey_point_input", uploaded_filename_default)
    st.session_state.setdefault("survey_date_text_input", "")
    st.session_state.setdefault("weather_input", DEFAULT_WEATHER)
    st.session_state.setdefault("responsible_party_input", DEFAULT_RESPONSIBLE_PARTY)
    st.session_state.setdefault("survey_period_input", DEFAULT_SURVEY_PERIOD)
    st.session_state.setdefault("north_label_input", "")
    st.session_state.setdefault("south_label_input", "")
    st.session_state.setdefault("east_label_input", "")
    st.session_state.setdefault("west_label_input", "")
    st.session_state.setdefault("north_road_input", "")
    st.session_state.setdefault("south_road_input", "")
    st.session_state.setdefault("east_road_input", "")
    st.session_state.setdefault("west_road_input", "")
    st.session_state.setdefault("caption_text_input", DEFAULT_CAPTION_TEXT)
    st.session_state.setdefault("peak_mode_select", DEFAULT_PEAK_MODE)
    for _peak_key, _peak_default in default_peak_window_state().items():
        st.session_state.setdefault(_peak_key, _peak_default)
    st.session_state.setdefault("show_u_turn_checkbox", True)
    st.session_state.setdefault("use_template_report_layout_checkbox", True)
    st.session_state.setdefault("use_excel_com_native_charts_checkbox", bool(excel_com_status.available))
    st.session_state.setdefault("mapping_editor_version", 0)
    _ensure_pce_factor_state()

    with st.sidebar:
        _render_project_session_section(
            uploaded_file.name if uploaded_file is not None else None,
            len(file_bytes) if uploaded_file is not None else None,
            compact=True,
        )
        st.divider()
        st.subheader("โหมดส่งออกรายงาน")
        if excel_com_status.available:
            export_mode_options = [EXCEL_TEMPLATE_EXPORT_MODE, SAFE_PNG_EXPORT_MODE]
            if st.session_state.get("report_export_mode_radio") not in export_mode_options:
                if st.session_state.get("use_excel_com_native_charts_checkbox", True):
                    st.session_state["report_export_mode_radio"] = EXCEL_TEMPLATE_EXPORT_MODE
                else:
                    st.session_state["report_export_mode_radio"] = SAFE_PNG_EXPORT_MODE
        else:
            export_mode_options = [SAFE_PNG_EXPORT_MODE]
            st.session_state["report_export_mode_radio"] = SAFE_PNG_EXPORT_MODE

        export_mode = st.radio("โหมดส่งออกรายงาน", options=export_mode_options, key="report_export_mode_radio")
        use_excel_com_native_charts = bool(excel_com_status.available and export_mode == EXCEL_TEMPLATE_EXPORT_MODE)
        st.session_state["use_excel_com_native_charts_checkbox"] = use_excel_com_native_charts
        st.caption(
            "ใช้ Excel Template เมื่อ COM พร้อมใช้งาน"
            if use_excel_com_native_charts
            else "โหมดสำรองใช้กราฟ PNG แบบคงที่"
        )

        st.divider()
        st.subheader("Excel Engine")
        previous_excel_com_available = bool(excel_com_status.available)
        if st.button("ทดสอบ Excel COM", key="test_excel_com"):
            excel_com_status = _probe_excel_com_for_ui(force=True)
            if excel_com_status.available and not previous_excel_com_available:
                st.session_state["report_export_mode_radio"] = EXCEL_TEMPLATE_EXPORT_MODE

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

    st.markdown(
        """
        <div class="tmc-header">
            <h1 class="tmc-title">TMC Processor</h1>
            <div class="tmc-subtitle">ประมวลผลข้อมูล Turning Movement Count และจัดทำรายงาน Excel</div>
            <div class="tmc-workflow">อัปโหลดไฟล์ → ตั้งค่างาน → กำหนดทิศทาง → ประมวลผล → ตรวจสอบช่วงเร่งด่วน → ส่งออกไฟล์</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    _render_status_cards(
        uploaded_name=uploaded_file.name if uploaded_file is not None else None,
        excel_com_status=excel_com_status,
        export_mode=export_mode,
    )

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

    setup_tab, mapping_tab, dashboard_tab, export_tab, qa_tab = st.tabs(
        ["ตั้งค่า", "กำหนดทิศทาง", "ตรวจ Peak", "ส่งออก", "QA / ขั้นสูง"]
    )

    with setup_tab:
        st.header("ตั้งค่างาน")
        with st.container(border=True):
            _render_section_header("ข้อมูลรายงาน", "ข้อมูลหลักที่จะใช้ในหน้าแรกของรายงาน")
            report_cols = st.columns(3)
            project_name = report_cols[0].text_input("ชื่อโครงการ", key="project_name_input")
            tmc_id = report_cols[1].text_input("TMC ID", key="tmc_id_input")
            tmc_title = report_cols[2].text_input("ชื่อจุดนับ", key="tmc_title_input")
            info_cols = st.columns(4)
            survey_point = info_cols[0].text_input("จุดสำรวจ", key="survey_point_input")
            survey_date_text = info_cols[1].text_input("วันที่สำรวจ", key="survey_date_text_input")
            weather = info_cols[2].text_input("สภาพอากาศ", key="weather_input")
            responsible_party = info_cols[3].text_input("ผู้รับผิดชอบ", key="responsible_party_input")

        with st.container(border=True):
            _render_section_header("ค่าเริ่มต้นช่วงเร่งด่วน", "กำหนดกรอบเวลาที่ใช้คัดเลือก Peak อัตโนมัติ")
            period_row_1 = st.columns(4)
            period_row_2 = st.columns(4)
            survey_period = period_row_1[0].text_input("ช่วงเวลาสำรวจ", key="survey_period_input")
            if st.session_state.get("peak_mode_select", DEFAULT_PEAK_MODE) not in PEAK_MODE_OPTIONS:
                st.session_state["peak_mode_select"] = DEFAULT_PEAK_MODE
            peak_mode = period_row_1[1].selectbox("รูปแบบการคำนวณ Peak", options=PEAK_MODE_OPTIONS, key="peak_mode_select")
            am_peak_window_start = period_row_1[2].time_input("เริ่มช่วง AM", step=900, key="am_peak_window_start_input")
            am_peak_window_end = period_row_1[3].time_input("สิ้นสุดช่วง AM", step=900, key="am_peak_window_end_input")
            pm_peak_window_start = period_row_2[0].time_input("เริ่มช่วง PM", step=900, key="pm_peak_window_start_input")
            pm_peak_window_end = period_row_2[1].time_input("สิ้นสุดช่วง PM", step=900, key="pm_peak_window_end_input")

        with st.container(border=True):
            _render_section_header("ป้ายปลายทาง", "ข้อความปลายทางของแต่ละขาเข้า-ออก")
            direction_cols = st.columns(4)
            north_label = direction_cols[0].text_input("ป้ายปลายทางด้านเหนือ", key="north_label_input")
            south_label = direction_cols[1].text_input("ป้ายปลายทางด้านใต้", key="south_label_input")
            east_label = direction_cols[2].text_input("ป้ายปลายทางด้านตะวันออก", key="east_label_input")
            west_label = direction_cols[3].text_input("ป้ายปลายทางด้านตะวันตก", key="west_label_input")

        with st.container(border=True):
            _render_section_header("ชื่อถนน / ทางหลวง", "ชื่อถนนที่จะแสดงในแผนภาพและรายงาน")
            road_cols = st.columns(4)
            north_road = road_cols[0].text_input("ชื่อถนนด้านเหนือ", key="north_road_input")
            south_road = road_cols[1].text_input("ชื่อถนนด้านใต้", key="south_road_input")
            east_road = road_cols[2].text_input("ชื่อถนนด้านตะวันออก", key="east_road_input")
            west_road = road_cols[3].text_input("ชื่อถนนด้านตะวันตก", key="west_road_input")
            st.markdown("#### คำบรรยายรูป Diagram")
            caption_text = st.text_input("คำบรรยายรูป Diagram", key="caption_text_input")
            show_u_turn = st.checkbox("แสดง movement กลับรถ", key="show_u_turn_checkbox")

        selected_pce_factors = _render_pce_factor_editor()

    setup = _setup_from_inputs(
        project_name=project_name,
        tmc_id=tmc_id,
        tmc_title=tmc_title,
        survey_point=survey_point,
        survey_date_text=survey_date_text,
        weather=weather,
        responsible_party=responsible_party,
        survey_period=survey_period,
        north_label=north_label,
        south_label=south_label,
        east_label=east_label,
        west_label=west_label,
        north_road=north_road,
        south_road=south_road,
        east_road=east_road,
        west_road=west_road,
        caption_text=caption_text,
        uploaded_filename=uploaded_file.name if uploaded_file is not None else "",
        peak_mode=peak_mode,
        am_peak_window_start=am_peak_window_start,
        am_peak_window_end=am_peak_window_end,
        pm_peak_window_start=pm_peak_window_start,
        pm_peak_window_end=pm_peak_window_end,
        show_u_turn=show_u_turn,
    )
    peak_windows = {
        "AM": (setup["am_peak_window_start"], setup["am_peak_window_end"]),
        "PM": (setup["pm_peak_window_start"], setup["pm_peak_window_end"]),
    }
    mapping = pd.DataFrame(st.session_state.get("mapping_table") or [])

    with mapping_tab:
        st.header("กำหนดทิศทาง")
        if uploaded_file is None:
            _render_empty_state(
                "ยังไม่มีไฟล์สำรวจ",
                "อัปโหลดไฟล์ Excel ในแถบด้านซ้ายเพื่อเริ่มกำหนดทิศทางจาก Sheet สำรวจ",
            )
        elif not detected_sheet_names:
            st.warning('ไม่พบ Sheet ทิศทางจากไฟล์สำรวจ ควรมีชื่อ Sheet เช่น "ทิศ 1", "ทิศ 2", หรือ "ทิศ 2+3"')
        else:
            st.markdown("#### Sheet ทิศทางที่ตรวจพบ")
            st.dataframe(preview_summary, width="stretch")
            _render_section_header(
                "ตารางกำหนดทิศทาง",
                "ตรวจสอบ Sheet สำรวจและกำหนด movement สำหรับรายงาน",
            )
            st.caption("หลายแถวสามารถ map ไปยัง movement เดียวกันได้ เช่น ทางหลักตรง + ทางคู่ขนานตรง → NS")
            default_mapping = default_mapping_for_sheets(detected_sheet_names)
            if st.session_state.get("mapping_table") is not None:
                default_mapping = apply_saved_mapping_to_sheets(detected_sheet_names, pd.DataFrame(st.session_state["mapping_table"]))

            st.markdown("#### นำเข้า/ส่งออก Mapping")
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
                    mapping_to_excel_bytes(default_mapping),
                    "tmc_mapping_template.xlsx",
                    EXCEL_MIME,
                )
            if mapping_upload is not None:
                try:
                    mapping_upload_bytes = mapping_upload.getvalue()
                    mapping_upload_identity = (mapping_upload.name, hashlib.sha256(mapping_upload_bytes).hexdigest())
                    if st.session_state.get("tmc_mapping_upload_identity") != mapping_upload_identity:
                        default_mapping = apply_saved_mapping_to_sheets(
                            detected_sheet_names,
                            read_mapping_excel(BytesIO(mapping_upload_bytes)),
                        )
                        st.session_state["mapping_table"] = default_mapping.to_dict("records")
                        st.session_state["mapping_editor_version"] = int(st.session_state.get("mapping_editor_version", 0) or 0) + 1
                        st.session_state["tmc_mapping_table_from_session"] = False
                        st.session_state["tmc_mapping_upload_identity"] = mapping_upload_identity
                        st.success("โหลด Mapping และปรับใช้กับ Sheet ที่ตรวจพบแล้ว")
                except Exception as exc:  # pragma: no cover - UI guardrail
                    st.error(f"ไม่สามารถโหลดไฟล์ Mapping ได้: {exc}")

            preset_name_seed = st.session_state.get("tmc_id_input") or st.session_state.get("tmc_title_input") or uploaded_file.name
            preset_source = pd.DataFrame(st.session_state.get("mapping_table") or default_mapping.to_dict("records"))
            preset_bytes = serialize_mapping_preset(
                build_mapping_preset(
                    preset_source,
                    preset_name=str(preset_name_seed or "TMC Mapping Preset"),
                )
            )
            st.session_state["tmc_mapping_preset_bytes"] = preset_bytes
            st.session_state["tmc_mapping_preset_filename"] = safe_mapping_preset_filename(preset_name_seed)
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
            if mapping_preset_upload is not None:
                try:
                    preset_upload_bytes = mapping_preset_upload.getvalue()
                    preset_upload_identity = (
                        mapping_preset_upload.name,
                        hashlib.sha256(preset_upload_bytes).hexdigest(),
                    )
                    if st.session_state.get("tmc_mapping_preset_upload_identity") != preset_upload_identity:
                        loaded_preset = load_mapping_preset(preset_upload_bytes)
                        apply_result = apply_mapping_preset_to_detected_sheets(loaded_preset, detected_sheet_names)
                        default_mapping = apply_result.mapping
                        st.session_state["mapping_table"] = default_mapping.to_dict("records")
                        st.session_state["mapping_editor_version"] = int(st.session_state.get("mapping_editor_version", 0) or 0) + 1
                        st.session_state["tmc_mapping_table_from_session"] = False
                        st.session_state["tmc_mapping_preset_upload_identity"] = preset_upload_identity
                        st.session_state["tmc_mapping_preset_apply_info"] = {
                            "matched": apply_result.matched_sheet_count,
                            "missing": apply_result.missing_detected_sheet_count,
                            "extra": apply_result.extra_preset_row_count,
                        }
                        st.session_state["tmc_mapping_preset_warnings"] = list(loaded_preset.warnings)
                except (MappingPresetError, ValueError) as exc:
                    st.error(f"ไม่สามารถเปิด Mapping Preset ได้: {exc}")
            preset_info = st.session_state.get("tmc_mapping_preset_apply_info")
            if preset_info:
                st.success("โหลด Mapping Preset สำเร็จ")
                st.info(
                    f"{preset_info.get('matched', 0)} sheets matched; "
                    f"{preset_info.get('missing', 0)} detected sheets still need review; "
                    f"{preset_info.get('extra', 0)} preset rows were not found in current workbook."
                )
            for warning_message in st.session_state.get("tmc_mapping_preset_warnings", []):
                st.warning(warning_message)

            mapping_editor_version = int(st.session_state.get("mapping_editor_version", 0) or 0)
            for warning_message in mapping_control_warnings(default_mapping):
                st.warning(_thai_mapping_control_warning(warning_message))
            movement_code_options = selectbox_options_with_existing_values(
                ["", *MOVEMENT_CODE_OPTIONS],
                default_mapping["movement_code"] if "movement_code" in default_mapping else None,
            )
            mapping = st.data_editor(
                _ordered_mapping_frame(default_mapping),
                width="stretch",
                num_rows="dynamic",
                column_config={
                    "raw_sheet": st.column_config.TextColumn("Sheet ต้นทาง", disabled=True),
                    "raw_direction": st.column_config.TextColumn("ทิศทางต้นทาง", disabled=True),
                    "movement_code": st.column_config.SelectboxColumn(
                        "movement_code สำหรับรายงาน",
                        options=movement_code_options,
                        required=True,
                    ),
                    "from_leg": st.column_config.SelectboxColumn("from_leg", options=["", *LEG_OPTIONS], required=True),
                    "to_leg": st.column_config.SelectboxColumn("to_leg", options=["", *LEG_OPTIONS], required=True),
                    "turn_type": st.column_config.SelectboxColumn("turn_type", options=["", *TURN_TYPE_OPTIONS], required=True),
                    "facility_type": st.column_config.SelectboxColumn(
                        "facility_type",
                        options=["", *FACILITY_TYPE_OPTIONS],
                        required=True,
                        default="at_grade",
                    ),
                    "source_stream": st.column_config.SelectboxColumn(
                        "source_stream",
                        options=SOURCE_STREAM_OPTIONS,
                        required=False,
                        default="mainline",
                    ),
                    "raw_movement_label": st.column_config.TextColumn("ป้าย movement ต้นทาง", required=False),
                    "include_in_peak": st.column_config.CheckboxColumn("ใช้คำนวณ Peak", default=True),
                    "include_in_report": st.column_config.CheckboxColumn("แสดงในรายงาน", default=True),
                    "aggregation_method": st.column_config.SelectboxColumn(
                        "aggregation_method",
                        options=AGGREGATION_METHOD_OPTIONS,
                        required=False,
                        default="sum",
                    ),
                },
                key=f"mapping_editor_{mapping_editor_version}",
            )
            mapping = pd.DataFrame(mapping)
            st.session_state["mapping_table"] = mapping.to_dict("records")
            st.session_state["tmc_mapping_table_from_session"] = False

            st.markdown("#### ตรวจสอบการรวม movement")
            for aggregation_message in movement_aggregation_messages(mapping):
                st.info(_thai_aggregation_message(aggregation_message))

            mapping_issues = validate_mapping_for_processing(detected_sheet_names, mapping)
            if mapping_issues.empty:
                st.success("การกำหนดทิศทางพร้อมสำหรับประมวลผล")
            else:
                st.warning("กรุณาตรวจสอบ Mapping ก่อนประมวลผล")
                st.dataframe(_mapping_issue_display(mapping_issues), width="stretch")

            run = st.button("ประมวลผลไฟล์ TMC", type="primary", disabled=not mapping_issues.empty)
            if run:
                try:
                    raw_sheets = {name: parsed.data for name, parsed in parsed_details.items()}
                    result = process_tmc(
                        raw_sheets=raw_sheets,
                        mapping=mapping,
                        setup=setup,
                        detected_sheets=detected_sheet_names,
                        peak_mode=peak_mode,
                        peak_windows=peak_windows,
                        pce_factors=selected_pce_factors,
                        generate_workbook=False,
                    )
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
                    st.success("ประมวลผลเสร็จแล้ว กรุณาตรวจสอบช่วงเร่งด่วนในแท็บ “ตรวจสอบกราฟและช่วงเร่งด่วน”")

    processed = st.session_state.get("tmc_processed")
    pce_results_stale = _processed_pce_results_stale(processed, selected_pce_factors)
    if pce_results_stale:
        st.session_state["tmc_pce_results_stale"] = True
        st.session_state.pop("tmc_output", None)
    result = None if pce_results_stale else (processed["result"] if processed else None)
    mapping_df = processed["mapping"] if processed and not pce_results_stale else mapping
    hourly_movement = hourly_movement_pcu(result.normalized, mapping_df) if result is not None else pd.DataFrame()
    confirmed_am_start = st.session_state.get("tmc_confirmed_am_peak_start", "")
    confirmed_am_end = st.session_state.get("tmc_confirmed_am_peak_end", "")
    confirmed_pm_start = st.session_state.get("tmc_confirmed_pm_peak_start", "")
    confirmed_pm_end = st.session_state.get("tmc_confirmed_pm_peak_end", "")

    with dashboard_tab:
        st.header("ตรวจสอบกราฟและช่วงเร่งด่วน")
        if pce_results_stale:
            st.warning("ค่า PCE เปลี่ยนหลังจากประมวลผลแล้ว กรุณาประมวลผลใหม่ก่อนตรวจสอบกราฟหรือส่งออกรายงาน")
        if result is None:
            _render_empty_state(
                "ยังไม่มีผลประมวลผล",
                "ประมวลผลไฟล์ที่กำหนดทิศทางแล้ว เพื่อดูกราฟ PCU รายชั่วโมงและยืนยันช่วงเร่งด่วน",
            )
        else:
            st.markdown("#### สรุปผลการประมวลผล")
            metric_cols = st.columns(4)
            metric_cols[0].metric("จำนวนแถวข้อมูลที่ปรับรูปแบบแล้ว", f"{len(result.normalized):,}")
            metric_cols[1].metric("จำนวนรถรวม", f"{result.normalized['count'].sum():,.0f}" if not result.normalized.empty else "0")
            metric_cols[2].metric("PCU รวม", f"{result.normalized['pcu'].sum():,.0f}" if not result.normalized.empty else "0")
            metric_cols[3].metric("ประเด็นตรวจสอบข้อมูล", f"{len(result.qc):,}")

            _render_qc_status(result.qc)

            chart_frame = hourly_interval_rows(hourly_movement)
            st.markdown("#### ปริมาณจราจรรวมรายชั่วโมง")
            if not chart_frame.empty:
                time_column = chart_frame.columns[0]
                if "Total" in chart_frame:
                    st.line_chart(chart_frame.set_index(time_column)[["Total"]], width="stretch")
            else:
                _render_empty_state("ไม่มีข้อมูล PCU รายชั่วโมง", "ยังไม่มีข้อมูลที่เพียงพอสำหรับแสดงกราฟ")

            am_start, am_end, am_pcu = _peak_period_text(result.peaks, "AM")
            pm_start, pm_end, pm_pcu = _peak_period_text(result.peaks, "PM")
            interval_options = _hourly_interval_options(hourly_movement, result.peaks)
            if interval_options:
                option_labels = [label for label, _, _ in interval_options]
                am_default = f"{am_start}-{am_end}" if am_start and am_end else option_labels[0]
                pm_default = f"{pm_start}-{pm_end}" if pm_start and pm_end else option_labels[min(1, len(option_labels) - 1)]
                loaded_confirmed_peaks = st.session_state.get("tmc_loaded_confirmed_peaks") or {}
                loaded_am_label = f"{loaded_confirmed_peaks.get('am_peak_start', '')}-{loaded_confirmed_peaks.get('am_peak_end', '')}".strip("-")
                loaded_pm_label = f"{loaded_confirmed_peaks.get('pm_peak_start', '')}-{loaded_confirmed_peaks.get('pm_peak_end', '')}".strip("-")
                if loaded_am_label in option_labels and "am_peak_period_select" not in st.session_state:
                    st.session_state["am_peak_period_select"] = loaded_am_label
                    am_default = loaded_am_label
                if loaded_pm_label in option_labels and "pm_peak_period_select" not in st.session_state:
                    st.session_state["pm_peak_period_select"] = loaded_pm_label
                    pm_default = loaded_pm_label
                am_index = option_labels.index(am_default) if am_default in option_labels else 0
                pm_index = option_labels.index(pm_default) if pm_default in option_labels else min(1, len(option_labels) - 1)

                _render_section_header(
                    "ยืนยันช่วงเร่งด่วน",
                    "เปรียบเทียบค่าที่ระบบแนะนำกับช่วงเวลาที่ต้องการใช้ในรายงาน",
                )
                confirm_cols = st.columns(2)
                am_peak_label = confirm_cols[0].selectbox("เลือกช่วงเร่งด่วนเช้า", option_labels, index=am_index, key="am_peak_period_select")
                pm_peak_label = confirm_cols[1].selectbox("เลือกช่วงเร่งด่วนเย็น", option_labels, index=pm_index, key="pm_peak_period_select")
                confirmed_am_start, confirmed_am_end = _selected_interval(interval_options, am_peak_label)
                confirmed_pm_start, confirmed_pm_end = _selected_interval(interval_options, pm_peak_label)
                st.session_state["tmc_confirmed_am_peak_start"] = confirmed_am_start
                st.session_state["tmc_confirmed_am_peak_end"] = confirmed_am_end
                st.session_state["tmc_confirmed_pm_peak_start"] = confirmed_pm_start
                st.session_state["tmc_confirmed_pm_peak_end"] = confirmed_pm_end
                st.caption("รายงาน Excel จะใช้ช่วงเร่งด่วนที่ยืนยันในหน้านี้")

                peak_cols = st.columns(4)
                with peak_cols[0]:
                    _render_peak_card("ช่วงเร่งด่วนเช้า (แนะนำ)", f"{am_start}-{am_end}" if am_start and am_end else "", am_pcu, "auto_suggested")
                with peak_cols[1]:
                    _render_peak_card("ช่วงเร่งด่วนเย็น (แนะนำ)", f"{pm_start}-{pm_end}" if pm_start and pm_end else "", pm_pcu, "auto_suggested")
                with peak_cols[2]:
                    confirmed_am_label = f"{confirmed_am_start}-{confirmed_am_end}" if confirmed_am_start and confirmed_am_end else ""
                    _render_peak_card("ช่วงเร่งด่วนเช้า (ยืนยัน)", confirmed_am_label, _interval_total_pcu(hourly_movement, confirmed_am_label), "user_confirmed")
                with peak_cols[3]:
                    confirmed_pm_label = f"{confirmed_pm_start}-{confirmed_pm_end}" if confirmed_pm_start and confirmed_pm_end else ""
                    _render_peak_card("ช่วงเร่งด่วนเย็น (ยืนยัน)", confirmed_pm_label, _interval_total_pcu(hourly_movement, confirmed_pm_label), "user_confirmed")
            else:
                st.warning("ไม่มีช่วงเวลารายชั่วโมงสำหรับยืนยัน Peak")

            with st.expander("ตารางปริมาณจราจรแยกตามทิศทาง", expanded=False):
                st.dataframe(hourly_movement, width="stretch")

    with export_tab:
        st.header("ส่งออกไฟล์")
        st.markdown("#### โหมดส่งออกรายงาน")
        if export_mode == EXCEL_TEMPLATE_EXPORT_MODE:
            st.markdown('<div class="tmc-mode-note tmc-mode-note-success"><strong>Excel Template Mode</strong> · แนะนำสำหรับรายงานฉบับใช้งานจริง</div>', unsafe_allow_html=True)
            st.caption("รักษากราฟ Native Chart สูตร และรูปแบบเทมเพลต Excel เมื่อ Excel COM พร้อมใช้งาน")
        else:
            st.markdown('<div class="tmc-mode-note tmc-mode-note-warning"><strong>Safe PNG Export Mode</strong> · โหมดสำรอง</div>', unsafe_allow_html=True)
            st.caption("ใช้กราฟ PNG แบบคงที่ เหมาะเมื่อ Excel COM ใช้งานไม่ได้")

        st.markdown("#### Excel Engine")
        if st.button("ทดสอบ Excel COM", key="test_excel_com_export_tab"):
            excel_com_status = _probe_excel_com_for_ui(force=True)
            if excel_com_status.available:
                version_text = f" — Excel version: {excel_com_status.version}" if excel_com_status.version else ""
                st.success(f"Excel COM พร้อมใช้งาน{version_text}")
            else:
                st.warning(f"Excel COM ไม่พร้อมใช้งาน ระบบจะใช้โหมดสำรองแบบ PNG: {excel_com_status.reason}")
        with st.expander("รายละเอียด Excel COM", expanded=False):
            _render_excel_com_status(excel_com_status)

        _render_section_header("ความพร้อมก่อนส่งออก", "รายการตรวจสอบก่อนสร้างรายงาน Excel")
        if result is not None:
            _render_qc_status(result.qc)
        confirmed_ready = all([confirmed_am_start, confirmed_am_end, confirmed_pm_start, confirmed_pm_end])
        _render_readiness_checklist(
            [
                ("โหลดไฟล์สำรวจแล้ว", uploaded_file is not None, ""),
                ("Mapping พร้อมใช้งาน", bool(st.session_state.get("mapping_table")), ""),
                ("ประมวลผลแล้ว", result is not None, "ค่า PCE เปลี่ยน กรุณาประมวลผลใหม่" if pce_results_stale else ""),
                ("ยืนยันช่วงเร่งด่วนแล้ว", confirmed_ready, ""),
                (
                    "Excel COM พร้อมใช้งาน",
                    bool(excel_com_status.available) if export_mode == EXCEL_TEMPLATE_EXPORT_MODE else True,
                    "จำเป็นสำหรับ Excel Template Mode" if export_mode == EXCEL_TEMPLATE_EXPORT_MODE else "ไม่จำเป็นในโหมดสำรอง",
                ),
            ]
        )

        use_template_report_layout = st.checkbox(
            "ใช้รูปแบบรายงานจาก Excel Template หากพร้อมใช้งาน",
            key="use_template_report_layout_checkbox",
        )
        st.markdown("#### ดาวน์โหลดไฟล์")
        if pce_results_stale:
            st.warning("ผลลัพธ์เดิมไม่ตรงกับค่า PCE ปัจจุบัน ระบบปิดการส่งออกไว้จนกว่าจะประมวลผลใหม่")
        export_run = st.button("สร้างรายงาน Excel", type="primary", disabled=not (result is not None and confirmed_ready and not pce_results_stale))
        if export_run:
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
                "am_peak_start": confirmed_am_start,
                "am_peak_end": confirmed_am_end,
                "pm_peak_start": confirmed_pm_start,
                "pm_peak_end": confirmed_pm_end,
                "peak_selection_source": PEAK_SELECTION_USER_CONFIRMED,
            }
            confirmed_periods = {
                "AM": (confirmed_am_start, confirmed_am_end),
                "PM": (confirmed_pm_start, confirmed_pm_end),
            }
            export_generated_at = generated_timestamp_text()
            try:
                raw_sheets = {name: parsed.data for name, parsed in parsed_details.items()}
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
                    confirmed_hourly_movement = hourly_movement_pcu(confirmed_result.normalized, mapping_df)
                    chart_pngs = report_chart_pngs(
                        confirmed_hourly_movement,
                        vehicle_composition_report(confirmed_result.normalized),
                        setup=confirmed_setup,
                    )
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
                    st.success("สร้างรายงาน Excel เสร็จแล้ว")
            except Exception as exc:  # pragma: no cover - UI guardrail
                st.error(f"ส่งออกไฟล์ไม่สำเร็จ: {exc}")

        output = st.session_state.get("tmc_output")
        if output:
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
            package_bytes = create_export_package_zip(
                workbook_bytes=output["workbook_bytes"],
                workbook_filename=output["workbook_filename"],
                export_summary_text=summary_text,
                project_session_bytes=session_bytes,
                project_session_filename=session_filename,
                mapping_preset_bytes=serialize_mapping_preset(
                    build_mapping_preset(
                        mapping_df,
                        preset_name=str(st.session_state.get("tmc_id_input") or st.session_state.get("tmc_title_input") or "TMC Mapping Preset"),
                    )
                ),
                mapping_preset_filename="mapping_preset.mapping.json",
                mapping=mapping_df,
                chart_pngs=output.get("chart_pngs", {}),
                diagram_png=output.get("diagram_png"),
            )
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
                st.image(output["diagram_png"], caption="Four-leg TMC movement diagram")
                _render_download_button("ดาวน์โหลด Diagram movement (PNG)", output["diagram_png"], "tmc_movement_diagram.png", PNG_MIME)
        else:
            _render_empty_state(
                "ยังไม่มีไฟล์ส่งออก",
                "ยืนยันช่วงเร่งด่วน AM/PM แล้วสร้างรายงาน Excel เมื่อพร้อม",
            )

    with qa_tab:
        st.header("ตรวจสอบข้อมูล / ขั้นสูง")
        if uploaded_file is None:
            _render_empty_state(
                "ยังไม่มี Workbook สำหรับ QA",
                "อัปโหลด Workbook เพื่อดูรายละเอียดการอ่านไฟล์และตรวจสอบข้อมูล",
            )
        else:
            with st.expander("รายละเอียดการอ่านไฟล์", expanded=False):
                st.dataframe(preview_summary, width="stretch")
                for sheet_name, preview in previews.items():
                    st.markdown(f"**{sheet_name}**")
                    st.dataframe(preview, width="stretch")
            with st.expander("รายละเอียด Parser", expanded=False):
                for sheet_name, parsed in parsed_details.items():
                    debug = parsed.debug
                    st.markdown(f"**{sheet_name}**")
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

        if result is not None:
            st.markdown("#### QC summary")
            qc_counts = _qc_severity_counts(result.qc)
            qc_cols = st.columns(3)
            qc_cols[0].metric("error", f"{qc_counts['error']:,}")
            qc_cols[1].metric("warning", f"{qc_counts['warning']:,}")
            qc_cols[2].metric("info", f"{qc_counts['info']:,}")
            if qc_counts["error"]:
                st.error("QC errors found. Review details before export.")
            elif qc_counts["warning"]:
                st.warning("QC warnings found. Export is still available, but review the notes first.")
            elif qc_counts["info"]:
                st.info("QC info notes are available for review.")
            else:
                st.success("No QC issues found.")
            with st.expander("QC details", expanded=False):
                st.dataframe(result.qc, width="stretch")
            with st.expander("Hourly summary", expanded=False):
                st.dataframe(result.hourly, width="stretch")
            with st.expander("Peak PHF", expanded=False):
                st.dataframe(result.peaks, width="stretch")
            with st.expander("ตัวอย่างข้อมูล Normalized", expanded=False):
                st.dataframe(result.normalized.head(1000), width="stretch")
            with st.expander("Movement Aggregation Audit", expanded=False):
                st.caption("แสดง source stream ที่ถูกรวมเป็น movement_code สำหรับรายงาน")
                st.dataframe(movement_aggregation_audit(result.normalized, mapping_df), width="stretch")
        else:
            _render_empty_state(
                "ยังไม่มีผล QA จากการประมวลผล",
                "ประมวลผลข้อมูลเพื่อดู QC, Normalized rows และรายละเอียดการรวม movement",
            )

        with st.expander("รายละเอียดเทมเพลต", expanded=False):
            _render_template_audit_notes()
        with st.expander("รายละเอียด Excel COM", expanded=False):
            _render_excel_com_status(excel_com_status)


_run_streamlit_app()
st.stop()
