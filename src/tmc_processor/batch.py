"""Basic batch processing helpers for TMC workbooks."""

from __future__ import annotations

from dataclasses import dataclass, field
from io import BytesIO
from pathlib import Path
import re
import warnings
from typing import Any, Iterable
from zipfile import ZIP_DEFLATED, ZipFile

import pandas as pd

from .charts import report_chart_pngs
from .diagram import DiagramConfig, generate_four_leg_tmc_diagram
from .export_package import build_export_summary_text
from .importer import load_detected_sheets
from .mapping import clean_mapping
from .mapping_preset import (
    apply_mapping_preset_to_detected_sheets,
    build_mapping_preset,
    serialize_mapping_preset,
)
from .metadata import APP_VERSION, TEMPLATE_VERSION, generated_timestamp_text, setup_with_metadata
from .pipeline import ProcessingResult, process_tmc
from .session import build_project_session, session_to_json_bytes
from .summaries import hourly_movement_pcu, vehicle_composition_report
from .time_utils import hourly_interval_options


BATCH_PACKAGE_MIME = "application/zip"
SAFE_BATCH_EXPORT_MODE = "Safe PNG Export Mode - Batch v1"
BATCH_EXCEL_TEMPLATE_EXPORT_MODE = "Excel Template Mode"
BATCH_SAFE_PNG_EXPORT_MODE = "Safe PNG Export Mode"
BATCH_STALE_MESSAGE_TH = "ข้อมูล Batch มีการเปลี่ยนแปลง กรุณาวิเคราะห์ Batch ใหม่"
BATCH_SUMMARY_COLUMNS = [
    "file_name",
    "survey_date_text",
    "output_stem",
    "folder_name",
    "status",
    "export_mode_requested",
    "export_mode_used",
    "export_status",
    "export_error",
    "suggested_AM_peak",
    "suggested_PM_peak",
    "confirmed_AM_peak",
    "confirmed_PM_peak",
    "AM_peak",
    "PM_peak",
    "total_vehicles",
    "total_PCU",
    "QC_errors",
    "QC_warnings",
    "QC_info",
    "export_file",
    "generated_report_filename",
    "notes",
]
BATCH_QC_COLUMNS = [
    "file_name",
    "output_stem",
    "survey_date_text",
    "severity",
    "category",
    "check",
    "message",
    "detail",
    "affected_field",
    "movement_code",
    "raw_sheet",
    "notes",
]


@dataclass(frozen=True)
class BatchItem:
    """One uploaded workbook in a Basic Batch run."""

    file_name: str
    workbook_bytes: bytes
    survey_date_text: str = ""
    output_stem: str = ""
    notes: str = ""

    def __post_init__(self) -> None:
        if not self.survey_date_text:
            object.__setattr__(self, "survey_date_text", derive_survey_date_text_from_filename(self.file_name))
        if not self.output_stem:
            object.__setattr__(self, "output_stem", safe_output_stem(self.file_name))


@dataclass(frozen=True)
class BatchSummaryRow:
    file_name: str
    survey_date_text: str
    output_stem: str
    folder_name: str
    status: str
    export_mode_requested: str = ""
    export_mode_used: str = ""
    export_status: str = ""
    export_error: str = ""
    suggested_AM_peak: str = ""
    suggested_PM_peak: str = ""
    confirmed_AM_peak: str = ""
    confirmed_PM_peak: str = ""
    AM_peak: str = ""
    PM_peak: str = ""
    total_vehicles: float = 0.0
    total_PCU: float = 0.0
    QC_errors: int = 0
    QC_warnings: int = 0
    QC_info: int = 0
    export_file: str = ""
    generated_report_filename: str = ""
    notes: str = ""


@dataclass
class BatchResult:
    summary_rows: list[BatchSummaryRow] = field(default_factory=list)
    qc_rows: list[dict[str, str]] = field(default_factory=list)
    package_bytes: bytes = b""
    generated_at: str = ""

    @property
    def has_failures(self) -> bool:
        return any(row.status == "failed" for row in self.summary_rows)


@dataclass
class BatchAnalysisItem:
    """Per-file analysis output used for Batch v1.1 peak review."""

    file_name: str
    survey_date_text: str
    output_stem: str
    folder_name: str
    status: str
    workbook_bytes: bytes = field(default=b"", repr=False)
    mapping: pd.DataFrame = field(default_factory=pd.DataFrame, repr=False)
    suggested_AM_peak: str = ""
    suggested_PM_peak: str = ""
    confirmed_AM_peak: str = ""
    confirmed_PM_peak: str = ""
    hourly_period_options: list[str] = field(default_factory=list)
    hourly_movement_pcu: pd.DataFrame = field(default_factory=pd.DataFrame, repr=False)
    qc: pd.DataFrame = field(default_factory=pd.DataFrame, repr=False)
    total_vehicles: float = 0.0
    total_PCU: float = 0.0
    QC_errors: int = 0
    QC_warnings: int = 0
    QC_info: int = 0
    mapping_status: str = ""
    matched_sheet_count: int = 0
    missing_detected_sheet_count: int = 0
    extra_preset_row_count: int = 0
    detected_sheets: list[str] = field(default_factory=list)
    notes: str = ""


@dataclass
class BatchAnalysisResult:
    items: list[BatchAnalysisItem] = field(default_factory=list)
    generated_at: str = ""
    mapping_preset_name: str = ""

    @property
    def successful_items(self) -> list[BatchAnalysisItem]:
        return [item for item in self.items if item.status == "success"]

    @property
    def has_failures(self) -> bool:
        return any(item.status == "failed" for item in self.items)


@dataclass
class _BatchFileArtifacts:
    folder_name: str
    output_stem: str
    workbook_bytes: bytes
    export_summary_text: str
    session_bytes: bytes
    mapping_preset_bytes: bytes
    chart_pngs: dict[str, bytes] = field(default_factory=dict)
    diagram_png: bytes | None = None


def safe_batch_name(name: str | None, default: str = "file") -> str:
    """Return an ASCII ZIP-safe member segment without paths."""

    base = Path(str(name or default)).name
    stem = Path(base).stem if "." in base else base
    cleaned = re.sub(r"[^A-Za-z0-9._-]+", "_", stem).strip("._")
    return (cleaned or default)[:80]


def safe_output_stem(name: str | None, default: str = "file") -> str:
    """Return a safe user-editable output stem for ZIP file names."""

    return safe_batch_name(name, default)


def unique_safe_output_stems(items: Iterable[BatchItem]) -> list[str]:
    """Return deterministic, collision-free output stems for a batch."""

    stems: list[str] = []
    seen: dict[str, int] = {}
    for index, item in enumerate(items, start=1):
        fallback = safe_output_stem(item.file_name, f"file_{index:02d}")
        base = safe_output_stem(item.output_stem, fallback)
        if not base:
            base = fallback
        duplicate_count = seen.get(base, 0)
        seen[base] = duplicate_count + 1
        stems.append(base if duplicate_count == 0 else f"{base}_{duplicate_count + 1:02d}")
    return stems


def batch_folder_name(index: int, output_stem: str) -> str:
    """Return the predictable per-file ZIP folder name."""

    return f"file_{index:02d}_{safe_output_stem(output_stem, f'file_{index:02d}')}"


def derive_survey_date_text_from_filename(filename: str | None) -> str:
    """Extract a practical survey date from common filename patterns."""

    stem = Path(str(filename or "")).stem
    patterns = [
        (r"(?<!\d)(20\d{2})[-_. ]?(0[1-9]|1[0-2])[-_. ]?([0-2]\d|3[01])(?!\d)", "{0}-{1}-{2}"),
        (r"(?<!\d)([0-2]\d|3[01])[-_. ](0[1-9]|1[0-2])[-_. ](20\d{2})(?!\d)", "{2}-{1}-{0}"),
    ]
    for pattern, template in patterns:
        match = re.search(pattern, stem)
        if match:
            return template.format(*match.groups())
    return ""


def batch_file_metadata_defaults(file_names: Iterable[str]) -> list[dict[str, str]]:
    """Build default UI metadata rows without raw paths."""

    return [
        {
            "file_name": Path(str(file_name)).name,
            "survey_date_text": derive_survey_date_text_from_filename(file_name),
            "output_stem": safe_output_stem(file_name, f"file_{index:02d}"),
            "notes": "",
        }
        for index, file_name in enumerate(file_names, start=1)
    ]


def batch_package_filename(name: str | None = None) -> str:
    cleaned = safe_batch_name(name or "tmc_batch", "tmc_batch")
    return f"{cleaned}_batch_package.zip"


def batch_inputs_ready(
    *,
    uploaded_workbook_count: int,
    mapping_available: bool,
    pce_factors_ready: bool = True,
) -> bool:
    """Return whether Basic Batch v1 has the required inputs to start."""

    return uploaded_workbook_count > 0 and mapping_available and pce_factors_ready


def batch_change_invalidates(previous: object, current: object, has_analysis: bool) -> bool:
    """Return whether changed Batch inputs should stale prior analysis/export."""

    return previous is not None and previous != current and has_analysis


def batch_zip_generation_block_reason(
    *,
    has_successful_files: bool,
    peaks_ready: bool,
    batch_stale: bool,
) -> str:
    """Return a user-facing reason when Batch ZIP generation must be blocked."""

    if batch_stale:
        return BATCH_STALE_MESSAGE_TH
    if not has_successful_files:
        return "ยังไม่มีไฟล์ Batch ที่วิเคราะห์สำเร็จ"
    if not peaks_ready:
        return "กรุณาเลือก AM/PM confirmed peak ให้ครบทุกไฟล์ที่สำเร็จก่อน Generate Batch ZIP"
    return ""


def batch_zip_contents_preview(summary_rows: Iterable[BatchSummaryRow]) -> list[str]:
    """Return a compact expected ZIP content outline for UI display."""

    rows = list(summary_rows)
    success_folders = [row.folder_name for row in rows if row.status == "success"]
    preview = ["batch_summary.xlsx"]
    if success_folders:
        for row in rows:
            if row.status != "success":
                continue
            folder = row.folder_name
            stem = safe_output_stem(row.output_stem or row.folder_name, "report")
            preview.extend(
                [
                    f"{folder}/",
                    f"{folder}/{stem}_report.xlsx",
                    f"{folder}/{stem}_export_summary.txt",
                    f"{folder}/{stem}_session.tmcproj.json",
                    f"{folder}/{stem}.mapping.json",
                    f"{folder}/charts/",
                ]
            )
    else:
        preview.append("(no successful file folders)")
    return preview


def _time_text(value: Any) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    return text[:5] if len(text) >= 5 and text[2:3] == ":" else text


def _peak_text(peaks: pd.DataFrame, period: str) -> str:
    if peaks.empty or "period" not in peaks.columns:
        return ""
    rows = peaks[peaks["period"].astype(str).str.upper() == period.upper()]
    if rows.empty:
        return ""
    row = rows.iloc[0]
    start = _time_text(row.get("peak_start"))
    end = _time_text(row.get("peak_end"))
    return f"{start}-{end}" if start and end else ""


def _confirmed_periods_from_peaks(peaks: pd.DataFrame) -> dict[str, tuple[str, str]]:
    periods: dict[str, tuple[str, str]] = {}
    for period in ("AM", "PM"):
        if peaks.empty or "period" not in peaks.columns:
            continue
        rows = peaks[peaks["period"].astype(str).str.upper() == period]
        if rows.empty:
            continue
        row = rows.iloc[0]
        start = _time_text(row.get("peak_start"))
        end = _time_text(row.get("peak_end"))
        if start and end:
            periods[period] = (start, end)
    return periods


def _period_text_to_tuple(value: str | None) -> tuple[str, str] | None:
    text = str(value or "").strip()
    if "-" not in text:
        return None
    start, end = [part.strip() for part in text.split("-", 1)]
    if not start or not end:
        return None
    return start[:5], end[:5]


def _confirmed_periods_from_labels(am_peak: str | None, pm_peak: str | None) -> dict[str, tuple[str, str]]:
    periods: dict[str, tuple[str, str]] = {}
    am_period = _period_text_to_tuple(am_peak)
    pm_period = _period_text_to_tuple(pm_peak)
    if am_period:
        periods["AM"] = am_period
    if pm_period:
        periods["PM"] = pm_period
    return periods


def _base_export_mode_label(value: str | None) -> str:
    text = str(value or "").strip()
    if text.casefold().startswith("excel template mode"):
        return BATCH_EXCEL_TEMPLATE_EXPORT_MODE
    if text.casefold().startswith("safe png export mode"):
        return BATCH_SAFE_PNG_EXPORT_MODE
    return text


def _export_used_from_warnings(requested_mode: str, export_warnings: Iterable[warnings.WarningMessage]) -> tuple[str, str]:
    requested = _base_export_mode_label(requested_mode)
    messages = [str(warning.message) for warning in export_warnings]
    fallback_markers = (
        "Excel COM unavailable",
        "Excel COM native-chart export failed",
        "falling back to safe openpyxl export with PNG charts",
    )
    if requested == BATCH_EXCEL_TEMPLATE_EXPORT_MODE and any(
        marker in message for message in messages for marker in fallback_markers
    ):
        return BATCH_SAFE_PNG_EXPORT_MODE, "; ".join(messages)
    return requested, "; ".join(messages)


def _qc_counts(qc: pd.DataFrame) -> dict[str, int]:
    if qc.empty or "severity" not in qc.columns:
        return {"error": 0, "warning": 0, "info": 0}
    counts = qc["severity"].fillna("").astype(str).str.casefold().value_counts()
    return {
        "error": int(counts.get("error", 0)),
        "warning": int(counts.get("warning", 0)),
        "info": int(counts.get("info", 0)),
    }


def _batch_qc_rows_for_file(
    *,
    file_name: str,
    output_stem: str,
    survey_date_text: str,
    qc: pd.DataFrame,
    notes: str = "",
) -> list[dict[str, str]]:
    if qc.empty:
        return []
    rows: list[dict[str, str]] = []
    for raw_row in qc.to_dict("records"):
        row = {
            "file_name": Path(file_name).name,
            "output_stem": safe_output_stem(output_stem, "file"),
            "survey_date_text": str(survey_date_text or ""),
            "severity": str(raw_row.get("severity", "") or ""),
            "category": str(raw_row.get("category", "") or ""),
            "check": str(raw_row.get("check", "") or ""),
            "message": str(raw_row.get("message", "") or ""),
            "detail": str(raw_row.get("detail", "") or ""),
            "affected_field": str(raw_row.get("affected_field", "") or ""),
            "movement_code": str(raw_row.get("movement_code", "") or ""),
            "raw_sheet": str(raw_row.get("raw_sheet", "") or ""),
            "notes": str(notes or ""),
        }
        rows.append(row)
    return rows


def _batch_failure_qc_row(
    *,
    file_name: str,
    output_stem: str,
    survey_date_text: str,
    message: str,
    notes: str = "",
) -> dict[str, str]:
    return {
        "file_name": Path(file_name).name,
        "output_stem": safe_output_stem(output_stem, "file"),
        "survey_date_text": str(survey_date_text or ""),
        "severity": "error",
        "category": "batch_processing",
        "check": "processing_failed",
        "message": str(message or "Batch processing failed."),
        "detail": "",
        "affected_field": "",
        "movement_code": "",
        "raw_sheet": "",
        "notes": str(notes or ""),
    }


def batch_qc_frame(qc_rows: Iterable[dict[str, Any]] | None) -> pd.DataFrame:
    """Return Batch_QC rows in stable workbook/UI column order."""

    return pd.DataFrame(list(qc_rows or []), columns=BATCH_QC_COLUMNS)


def batch_analysis_qc_rows(analysis: BatchAnalysisResult | None) -> list[dict[str, str]]:
    """Return QC preview rows from analyzed Batch files."""

    if analysis is None:
        return []
    rows: list[dict[str, str]] = []
    for item in analysis.items:
        if item.status == "success":
            rows.extend(
                _batch_qc_rows_for_file(
                    file_name=item.file_name,
                    output_stem=item.output_stem,
                    survey_date_text=item.survey_date_text,
                    qc=item.qc,
                    notes=item.notes,
                )
            )
        else:
            rows.append(
                _batch_failure_qc_row(
                    file_name=item.file_name,
                    output_stem=item.output_stem,
                    survey_date_text=item.survey_date_text,
                    message=item.notes,
                    notes=item.notes,
                )
            )
    return rows


def batch_selected_file_preview(item: BatchAnalysisItem | BatchSummaryRow) -> dict[str, Any]:
    """Return compact selected-file values for Batch Peak review."""

    return {
        "file_name": item.file_name,
        "survey_date_text": item.survey_date_text,
        "output_stem": item.output_stem,
        "status": item.status,
        "total_vehicles": getattr(item, "total_vehicles", 0.0),
        "total_PCU": getattr(item, "total_PCU", 0.0),
        "QC_errors": getattr(item, "QC_errors", 0),
        "QC_warnings": getattr(item, "QC_warnings", 0),
        "QC_info": getattr(item, "QC_info", 0),
        "suggested_AM_peak": getattr(item, "suggested_AM_peak", ""),
        "suggested_PM_peak": getattr(item, "suggested_PM_peak", ""),
        "confirmed_AM_peak": getattr(item, "confirmed_AM_peak", ""),
        "confirmed_PM_peak": getattr(item, "confirmed_PM_peak", ""),
    }


def _batch_summary_workbook(
    rows: list[BatchSummaryRow],
    *,
    qc_rows: list[dict[str, str]] | None = None,
    generated_at: str,
    mapping_preset_name: str,
) -> bytes:
    summary = pd.DataFrame([row.__dict__ for row in rows], columns=BATCH_SUMMARY_COLUMNS)
    batch_qc = batch_qc_frame(qc_rows)
    metadata = pd.DataFrame(
        [
            {"field": "app_version", "value": APP_VERSION},
            {"field": "template_version", "value": TEMPLATE_VERSION},
            {"field": "generated_at", "value": generated_at},
            {"field": "mapping_preset_name", "value": mapping_preset_name},
        ]
    )
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        metadata.to_excel(writer, sheet_name="metadata", index=False)
        summary.to_excel(writer, sheet_name="Batch_Summary", index=False)
        batch_qc.to_excel(writer, sheet_name="Batch_QC", index=False)
        for worksheet in writer.book.worksheets:
            worksheet.freeze_panes = "A2"
            for column_cells in worksheet.columns:
                max_length = max(len(str(cell.value)) if cell.value is not None else 0 for cell in column_cells)
                worksheet.column_dimensions[column_cells[0].column_letter].width = min(max(max_length + 2, 12), 48)
    return output.getvalue()


def create_batch_package_zip(
    *,
    summary_rows: list[BatchSummaryRow],
    qc_rows: list[dict[str, str]] | None = None,
    file_artifacts: list[_BatchFileArtifacts],
    generated_at: str,
    mapping_preset_name: str = "",
) -> bytes:
    """Create a Basic Batch ZIP package without embedding raw inputs."""

    output = BytesIO()
    summary_bytes = _batch_summary_workbook(
        summary_rows,
        qc_rows=list(qc_rows or []),
        generated_at=generated_at,
        mapping_preset_name=mapping_preset_name,
    )
    with ZipFile(output, "w", compression=ZIP_DEFLATED) as archive:
        archive.writestr("batch_summary.xlsx", summary_bytes)
        for artifact in file_artifacts:
            folder = safe_batch_name(artifact.folder_name, "file")
            stem = safe_output_stem(artifact.output_stem, folder)
            archive.writestr(f"{folder}/{stem}_report.xlsx", bytes(artifact.workbook_bytes))
            archive.writestr(f"{folder}/{stem}_export_summary.txt", artifact.export_summary_text.encode("utf-8"))
            archive.writestr(f"{folder}/{stem}_session.tmcproj.json", bytes(artifact.session_bytes))
            archive.writestr(f"{folder}/{stem}.mapping.json", bytes(artifact.mapping_preset_bytes))
            for chart_name, png_bytes in sorted(artifact.chart_pngs.items()):
                if png_bytes:
                    archive.writestr(f"{folder}/charts/{safe_batch_name(chart_name, 'chart')}.png", bytes(png_bytes))
            if artifact.diagram_png:
                archive.writestr(f"{folder}/charts/tmc_movement_diagram.png", bytes(artifact.diagram_png))
    return output.getvalue()


def _diagram_config(setup: dict[str, Any]) -> DiagramConfig:
    setup = setup_with_metadata(setup)
    return DiagramConfig(
        tmc_id=str(setup.get("tmc_id", "") or ""),
        tmc_name=str(setup.get("tmc_title") or setup.get("tmc_name", "") or ""),
        survey_date_text=str(setup.get("survey_date_text", "") or ""),
        north_label=str(setup.get("north_label", "") or ""),
        south_label=str(setup.get("south_label", "") or ""),
        east_label=str(setup.get("east_label", "") or ""),
        west_label=str(setup.get("west_label", "") or ""),
        north_road=str(setup.get("north_road", "") or ""),
        south_road=str(setup.get("south_road", "") or ""),
        east_road=str(setup.get("east_road", "") or ""),
        west_road=str(setup.get("west_road", "") or ""),
        survey_period_text=str(setup.get("survey_period") or setup.get("survey_period_text") or ""),
        caption_text=str(setup.get("caption_text", "") or ""),
        show_u_turn=bool(setup.get("show_u_turn", True)),
    )


def _process_one_file(
    item: BatchItem,
    *,
    folder_name: str,
    mapping: pd.DataFrame,
    setup: dict[str, Any],
    pce_factors: dict[str, float] | None,
    peak_mode: str,
    peak_windows: dict[str, tuple[str, str]] | None,
    export_mode: str,
    generated_at: str,
    use_template_report_layout: bool,
    use_excel_com_native_charts: bool,
    confirmed_peak_periods: dict[str, tuple[str, str]] | None = None,
    suggested_am_peak: str = "",
    suggested_pm_peak: str = "",
) -> tuple[BatchSummaryRow, _BatchFileArtifacts, list[dict[str, str]]]:
    output_stem = safe_output_stem(item.output_stem or item.file_name, folder_name)
    per_file_setup = {
        **setup,
        "survey_date_text": item.survey_date_text or str(setup.get("survey_date_text", "") or ""),
    }
    raw_sheets = load_detected_sheets(BytesIO(item.workbook_bytes))
    detected_sheets = list(raw_sheets)
    apply_result = apply_mapping_preset_to_detected_sheets(
        build_mapping_preset(mapping, preset_name="Batch Mapping Preset"),
        detected_sheets,
    )
    active_mapping = apply_result.mapping

    if confirmed_peak_periods is None:
        suggested = process_tmc(
            raw_sheets=raw_sheets,
            mapping=active_mapping,
            setup=per_file_setup,
            detected_sheets=detected_sheets,
            peak_mode=peak_mode,
            peak_windows=peak_windows,
            pce_factors=pce_factors,
            generate_workbook=False,
        )
        confirmed_periods = _confirmed_periods_from_peaks(suggested.peaks)
        suggested_am_peak = suggested_am_peak or _peak_text(suggested.peaks, "AM")
        suggested_pm_peak = suggested_pm_peak or _peak_text(suggested.peaks, "PM")
    else:
        confirmed_periods = confirmed_peak_periods
    confirmed_setup = {
        **per_file_setup,
        "peak_selection_source": "user_confirmed_batch",
    }
    if "AM" in confirmed_periods:
        confirmed_setup["am_peak_start"], confirmed_setup["am_peak_end"] = confirmed_periods["AM"]
    if "PM" in confirmed_periods:
        confirmed_setup["pm_peak_start"], confirmed_setup["pm_peak_end"] = confirmed_periods["PM"]

    with warnings.catch_warnings(record=True) as export_warnings:
        warnings.simplefilter("always", RuntimeWarning)
        result: ProcessingResult = process_tmc(
            raw_sheets=raw_sheets,
            mapping=active_mapping,
            setup=confirmed_setup,
            detected_sheets=detected_sheets,
            peak_mode=peak_mode,
            peak_windows=peak_windows,
            confirmed_peak_periods=confirmed_periods,
            pce_factors=pce_factors,
            generate_workbook=True,
            use_template_report_layout=use_template_report_layout,
            use_excel_com_native_charts=use_excel_com_native_charts,
            export_mode=export_mode,
            source_file_name=item.file_name,
            generated_at=generated_at,
        )
    export_mode_requested = _base_export_mode_label(export_mode)
    if export_mode_requested == BATCH_EXCEL_TEMPLATE_EXPORT_MODE and not use_excel_com_native_charts:
        export_mode_used = BATCH_SAFE_PNG_EXPORT_MODE
        export_warning_text = "Excel COM native chart export was not enabled; used Safe PNG Export Mode."
    else:
        export_mode_used, export_warning_text = _export_used_from_warnings(export_mode, export_warnings)
    hourly_movement = hourly_movement_pcu(result.normalized, active_mapping)
    chart_pngs = dict(
        report_chart_pngs(
            hourly_movement,
            vehicle_composition_report(result.normalized),
            setup=confirmed_setup,
        )
    )
    diagram_png = generate_four_leg_tmc_diagram(hourly_movement, result.peaks, _diagram_config(confirmed_setup))
    mapping_preset_bytes = serialize_mapping_preset(
        build_mapping_preset(active_mapping, preset_name="Batch Mapping Preset")
    )
    session = build_project_session(
        metadata=confirmed_setup,
        directions=confirmed_setup,
        mapping=active_mapping,
        detected_sheet_names=detected_sheets,
        peak_settings=confirmed_setup,
        export_settings={
            "use_template_report_layout": use_template_report_layout,
            "use_excel_com_native_charts": use_excel_com_native_charts,
            "template_version": TEMPLATE_VERSION,
            "export_mode_requested": export_mode_requested,
            "export_mode_used": export_mode_used,
        },
        pce_factors=result.pce_factors,
        source_file_name=Path(item.file_name).name,
        source_file_size=len(item.workbook_bytes),
    )
    session_bytes = session_to_json_bytes(session)
    summary_text = build_export_summary_text(
        setup=confirmed_setup,
        source_file_name=item.file_name,
        export_mode=export_mode,
        peaks=result.peaks,
        mapping=active_mapping,
        qc=result.qc,
        workbook_filename=f"{output_stem}_report.xlsx",
        pce_factors=result.pce_factors,
        export_settings={
            "template_version": TEMPLATE_VERSION,
            "export_mode_requested": export_mode_requested,
            "export_mode_used": export_mode_used,
            "export_status": "success",
            "export_error": "",
        },
        generated_at=generated_at,
    )
    summary_text = "\n".join(
        [
            summary_text.rstrip(),
            f"survey_date_text: {confirmed_setup.get('survey_date_text', '')}",
            f"output_stem: {output_stem}",
            f"export_mode_requested: {export_mode_requested}",
            f"export_mode_used: {export_mode_used}",
            "export_status: success",
            "export_error: ",
            "",
        ]
    )
    counts = _qc_counts(result.qc)
    row = BatchSummaryRow(
        file_name=Path(item.file_name).name,
        survey_date_text=item.survey_date_text or str(setup.get("survey_date_text", "") or ""),
        output_stem=output_stem,
        folder_name=folder_name,
        status="success",
        export_mode_requested=export_mode_requested,
        export_mode_used=export_mode_used,
        export_status="success",
        export_error="",
        suggested_AM_peak=suggested_am_peak,
        suggested_PM_peak=suggested_pm_peak,
        confirmed_AM_peak=_peak_text(result.peaks, "AM"),
        confirmed_PM_peak=_peak_text(result.peaks, "PM"),
        AM_peak=_peak_text(result.peaks, "AM"),
        PM_peak=_peak_text(result.peaks, "PM"),
        total_vehicles=float(result.normalized["count"].sum()) if "count" in result.normalized else 0.0,
        total_PCU=float(result.normalized["pcu"].sum()) if "pcu" in result.normalized else 0.0,
        QC_errors=counts["error"],
        QC_warnings=counts["warning"],
        QC_info=counts["info"],
        export_file=f"{folder_name}/{output_stem}_report.xlsx",
        generated_report_filename=f"{output_stem}_report.xlsx",
        notes=item.notes or export_warning_text or "Auto/suggested peaks confirmed by Batch v1.",
    )
    artifact = _BatchFileArtifacts(
        folder_name=folder_name,
        output_stem=output_stem,
        workbook_bytes=result.workbook_bytes,
        export_summary_text=summary_text,
        session_bytes=session_bytes,
        mapping_preset_bytes=mapping_preset_bytes,
        chart_pngs=chart_pngs,
        diagram_png=diagram_png,
    )
    qc_rows = _batch_qc_rows_for_file(
        file_name=item.file_name,
        output_stem=output_stem,
        survey_date_text=row.survey_date_text,
        qc=result.qc,
        notes=item.notes,
    )
    return row, artifact, qc_rows


def analyze_batch_files(
    items: Iterable[BatchItem],
    *,
    mapping: pd.DataFrame | None = None,
    mapping_preset: dict[str, Any] | None = None,
    setup: dict[str, Any] | None = None,
    pce_factors: dict[str, float] | None = None,
    peak_mode: str = "rolling_60min",
    peak_windows: dict[str, tuple[str, str]] | None = None,
    mapping_preset_name: str = "",
    generated_at: str | None = None,
) -> BatchAnalysisResult:
    """Analyze batch files enough for per-file peak confirmation."""

    generated_at = generated_at or generated_timestamp_text()
    setup = dict(setup or {})
    source_mapping = clean_mapping(mapping if mapping is not None else pd.DataFrame())
    if source_mapping.empty and mapping_preset:
        source_mapping = apply_mapping_preset_to_detected_sheets(mapping_preset, []).mapping

    analysis_items: list[BatchAnalysisItem] = []
    item_list = list(items)
    output_stems = unique_safe_output_stems(item_list)
    for index, item in enumerate(item_list, start=1):
        output_stem = output_stems[index - 1]
        folder_name = batch_folder_name(index, output_stem)
        per_file_setup = {
            **setup,
            "survey_date_text": item.survey_date_text or str(setup.get("survey_date_text", "") or ""),
        }
        try:
            raw_sheets = load_detected_sheets(BytesIO(item.workbook_bytes))
            detected_sheets = list(raw_sheets)
            active_mapping = (
                apply_mapping_preset_to_detected_sheets(mapping_preset, detected_sheets).mapping
                if source_mapping.empty and mapping_preset
                else source_mapping
            )
            apply_result = apply_mapping_preset_to_detected_sheets(
                build_mapping_preset(active_mapping, preset_name="Batch Mapping Preset"),
                detected_sheets,
            )
            active_mapping = apply_result.mapping
            result = process_tmc(
                raw_sheets=raw_sheets,
                mapping=active_mapping,
                setup=per_file_setup,
                detected_sheets=detected_sheets,
                peak_mode=peak_mode,
                peak_windows=peak_windows,
                pce_factors=pce_factors,
                generate_workbook=False,
            )
            counts = _qc_counts(result.qc)
            hourly_movement = hourly_movement_pcu(result.normalized, active_mapping)
            options = [option[0] for option in hourly_interval_options(hourly_movement)]
            suggested_am = _peak_text(result.peaks, "AM")
            suggested_pm = _peak_text(result.peaks, "PM")
            for suggested in (suggested_am, suggested_pm):
                if suggested and suggested not in options:
                    options.append(suggested)
            analysis_items.append(
                BatchAnalysisItem(
                    file_name=Path(item.file_name).name,
                    survey_date_text=item.survey_date_text or str(setup.get("survey_date_text", "") or ""),
                    output_stem=output_stem,
                    folder_name=folder_name,
                    status="success",
                    workbook_bytes=item.workbook_bytes,
                    mapping=active_mapping,
                    suggested_AM_peak=suggested_am,
                    suggested_PM_peak=suggested_pm,
                    confirmed_AM_peak=suggested_am,
                    confirmed_PM_peak=suggested_pm,
                    hourly_period_options=options,
                    hourly_movement_pcu=hourly_movement,
                    qc=result.qc,
                    total_vehicles=float(result.normalized["count"].sum()) if "count" in result.normalized else 0.0,
                    total_PCU=float(result.normalized["pcu"].sum()) if "pcu" in result.normalized else 0.0,
                    QC_errors=counts["error"],
                    QC_warnings=counts["warning"],
                    QC_info=counts["info"],
                    mapping_status="matched" if apply_result.missing_detected_sheet_count == 0 else "needs review",
                    matched_sheet_count=apply_result.matched_sheet_count,
                    missing_detected_sheet_count=apply_result.missing_detected_sheet_count,
                    extra_preset_row_count=apply_result.extra_preset_row_count,
                    detected_sheets=detected_sheets,
                    notes=item.notes or "Suggested peaks are ready for review.",
                )
            )
        except Exception as exc:
            analysis_items.append(
                BatchAnalysisItem(
                    file_name=Path(item.file_name).name,
                    survey_date_text=item.survey_date_text or str(setup.get("survey_date_text", "") or ""),
                    output_stem=output_stem,
                    folder_name=folder_name,
                    status="failed",
                    mapping_status="failed",
                    notes=item.notes or str(exc),
                )
            )
    return BatchAnalysisResult(
        items=analysis_items,
        generated_at=generated_at,
        mapping_preset_name=mapping_preset_name,
    )


def reviewed_peak_values_complete(analysis: BatchAnalysisResult) -> bool:
    """Return whether all successful analyzed files have confirmed AM/PM peaks."""

    return all(item.confirmed_AM_peak and item.confirmed_PM_peak for item in analysis.successful_items)


def generate_batch_zip_from_reviewed_peaks(
    analysis: BatchAnalysisResult,
    *,
    setup: dict[str, Any] | None = None,
    pce_factors: dict[str, float] | None = None,
    peak_mode: str = "rolling_60min",
    peak_windows: dict[str, tuple[str, str]] | None = None,
    export_mode: str = SAFE_BATCH_EXPORT_MODE,
    use_template_report_layout: bool = True,
    use_excel_com_native_charts: bool = False,
) -> BatchResult:
    """Generate the final Batch ZIP using reviewed per-file peak selections."""

    rows: list[BatchSummaryRow] = []
    qc_rows: list[dict[str, str]] = []
    artifacts: list[_BatchFileArtifacts] = []
    setup = dict(setup or {})
    export_mode_requested = _base_export_mode_label(export_mode)
    for item in analysis.items:
        if item.status != "success":
            rows.append(
                BatchSummaryRow(
                    file_name=Path(item.file_name).name,
                    survey_date_text=item.survey_date_text,
                    output_stem=item.output_stem,
                    folder_name=item.folder_name,
                    status="failed",
                    export_mode_requested=export_mode_requested,
                    export_mode_used="",
                    export_status="failed",
                    export_error=item.notes,
                    notes=item.notes,
                )
            )
            qc_rows.append(
                _batch_failure_qc_row(
                    file_name=item.file_name,
                    output_stem=item.output_stem,
                    survey_date_text=item.survey_date_text,
                    message=item.notes,
                    notes=item.notes,
                )
            )
            continue

        confirmed_periods = _confirmed_periods_from_labels(item.confirmed_AM_peak, item.confirmed_PM_peak)
        if "AM" not in confirmed_periods or "PM" not in confirmed_periods:
            rows.append(
                BatchSummaryRow(
                    file_name=Path(item.file_name).name,
                    survey_date_text=item.survey_date_text,
                    output_stem=item.output_stem,
                    folder_name=item.folder_name,
                    status="failed",
                    export_mode_requested=export_mode_requested,
                    export_mode_used="",
                    export_status="failed",
                    export_error="Confirmed AM/PM peak is missing.",
                    suggested_AM_peak=item.suggested_AM_peak,
                    suggested_PM_peak=item.suggested_PM_peak,
                    notes="Confirmed AM/PM peak is missing.",
                )
            )
            qc_rows.append(
                _batch_failure_qc_row(
                    file_name=item.file_name,
                    output_stem=item.output_stem,
                    survey_date_text=item.survey_date_text,
                    message="Confirmed AM/PM peak is missing.",
                    notes="Confirmed AM/PM peak is missing.",
                )
            )
            continue

        try:
            row, artifact, file_qc_rows = _process_one_file(
                BatchItem(
                    file_name=item.file_name,
                    workbook_bytes=item.workbook_bytes,
                    survey_date_text=item.survey_date_text,
                    output_stem=item.output_stem,
                    notes=item.notes,
                ),
                folder_name=item.folder_name,
                mapping=item.mapping,
                setup=setup,
                pce_factors=pce_factors,
                peak_mode=peak_mode,
                peak_windows=peak_windows,
                export_mode=export_mode,
                generated_at=analysis.generated_at,
                use_template_report_layout=use_template_report_layout,
                use_excel_com_native_charts=use_excel_com_native_charts,
                confirmed_peak_periods=confirmed_periods,
                suggested_am_peak=item.suggested_AM_peak,
                suggested_pm_peak=item.suggested_PM_peak,
            )
            rows.append(row)
            artifacts.append(artifact)
            qc_rows.extend(file_qc_rows)
        except Exception as exc:
            rows.append(
                BatchSummaryRow(
                    file_name=Path(item.file_name).name,
                    survey_date_text=item.survey_date_text,
                    output_stem=item.output_stem,
                    folder_name=item.folder_name,
                    status="failed",
                    export_mode_requested=export_mode_requested,
                    export_mode_used="",
                    export_status="failed",
                    export_error=str(exc),
                    suggested_AM_peak=item.suggested_AM_peak,
                    suggested_PM_peak=item.suggested_PM_peak,
                    confirmed_AM_peak=item.confirmed_AM_peak,
                    confirmed_PM_peak=item.confirmed_PM_peak,
                    notes=str(exc),
                )
            )
            qc_rows.append(
                _batch_failure_qc_row(
                    file_name=item.file_name,
                    output_stem=item.output_stem,
                    survey_date_text=item.survey_date_text,
                    message=str(exc),
                    notes=str(exc),
                )
            )

    package = create_batch_package_zip(
        summary_rows=rows,
        qc_rows=qc_rows,
        file_artifacts=artifacts,
        generated_at=analysis.generated_at,
        mapping_preset_name=analysis.mapping_preset_name,
    )
    return BatchResult(summary_rows=rows, qc_rows=qc_rows, package_bytes=package, generated_at=analysis.generated_at)


def process_batch_files(
    items: Iterable[BatchItem],
    *,
    mapping: pd.DataFrame | None = None,
    mapping_preset: dict[str, Any] | None = None,
    setup: dict[str, Any] | None = None,
    pce_factors: dict[str, float] | None = None,
    peak_mode: str = "rolling_60min",
    peak_windows: dict[str, tuple[str, str]] | None = None,
    export_mode: str = SAFE_BATCH_EXPORT_MODE,
    use_template_report_layout: bool = True,
    use_excel_com_native_charts: bool = False,
    mapping_preset_name: str = "",
    generated_at: str | None = None,
) -> BatchResult:
    """Process many uploaded workbooks with one shared mapping and settings."""

    generated_at = generated_at or generated_timestamp_text()
    analysis = analyze_batch_files(
        items,
        mapping=mapping,
        mapping_preset=mapping_preset,
        setup=setup,
        pce_factors=pce_factors,
        peak_mode=peak_mode,
        peak_windows=peak_windows,
        mapping_preset_name=mapping_preset_name,
        generated_at=generated_at,
    )
    return generate_batch_zip_from_reviewed_peaks(
        analysis,
        setup=setup,
        pce_factors=pce_factors,
        peak_mode=peak_mode,
        peak_windows=peak_windows,
        export_mode=export_mode,
        use_template_report_layout=use_template_report_layout,
        use_excel_com_native_charts=use_excel_com_native_charts,
    )
