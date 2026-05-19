"""Basic batch processing helpers for TMC workbooks."""

from __future__ import annotations

from dataclasses import dataclass, field
from io import BytesIO
from pathlib import Path
import re
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
BATCH_SUMMARY_COLUMNS = [
    "file_name",
    "folder_name",
    "status",
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
    "notes",
]


@dataclass(frozen=True)
class BatchItem:
    """One uploaded workbook in a Basic Batch run."""

    file_name: str
    workbook_bytes: bytes


@dataclass(frozen=True)
class BatchSummaryRow:
    file_name: str
    folder_name: str
    status: str
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
    notes: str = ""


@dataclass
class BatchResult:
    summary_rows: list[BatchSummaryRow] = field(default_factory=list)
    package_bytes: bytes = b""
    generated_at: str = ""

    @property
    def has_failures(self) -> bool:
        return any(row.status == "failed" for row in self.summary_rows)


@dataclass
class BatchAnalysisItem:
    """Per-file analysis output used for Batch v1.1 peak review."""

    file_name: str
    folder_name: str
    status: str
    workbook_bytes: bytes = field(default=b"", repr=False)
    mapping: pd.DataFrame = field(default_factory=pd.DataFrame, repr=False)
    suggested_AM_peak: str = ""
    suggested_PM_peak: str = ""
    confirmed_AM_peak: str = ""
    confirmed_PM_peak: str = ""
    hourly_period_options: list[str] = field(default_factory=list)
    total_vehicles: float = 0.0
    total_PCU: float = 0.0
    QC_errors: int = 0
    QC_warnings: int = 0
    QC_info: int = 0
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


def batch_zip_contents_preview(summary_rows: Iterable[BatchSummaryRow]) -> list[str]:
    """Return a compact expected ZIP content outline for UI display."""

    rows = list(summary_rows)
    success_folders = [row.folder_name for row in rows if row.status == "success"]
    preview = ["batch_summary.xlsx"]
    if success_folders:
        for folder in success_folders:
            preview.extend(
                [
                    f"{folder}/",
                    f"{folder}/report.xlsx",
                    f"{folder}/export_summary.txt",
                    f"{folder}/session.tmcproj.json",
                    f"{folder}/mapping_preset.mapping.json",
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


def _qc_counts(qc: pd.DataFrame) -> dict[str, int]:
    if qc.empty or "severity" not in qc.columns:
        return {"error": 0, "warning": 0, "info": 0}
    counts = qc["severity"].fillna("").astype(str).str.casefold().value_counts()
    return {
        "error": int(counts.get("error", 0)),
        "warning": int(counts.get("warning", 0)),
        "info": int(counts.get("info", 0)),
    }


def _batch_summary_workbook(
    rows: list[BatchSummaryRow],
    *,
    generated_at: str,
    mapping_preset_name: str,
) -> bytes:
    summary = pd.DataFrame([row.__dict__ for row in rows], columns=BATCH_SUMMARY_COLUMNS)
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
        summary.to_excel(writer, sheet_name="batch_summary", index=False)
        for worksheet in writer.book.worksheets:
            worksheet.freeze_panes = "A2"
            for column_cells in worksheet.columns:
                max_length = max(len(str(cell.value)) if cell.value is not None else 0 for cell in column_cells)
                worksheet.column_dimensions[column_cells[0].column_letter].width = min(max(max_length + 2, 12), 48)
    return output.getvalue()


def create_batch_package_zip(
    *,
    summary_rows: list[BatchSummaryRow],
    file_artifacts: list[_BatchFileArtifacts],
    generated_at: str,
    mapping_preset_name: str = "",
) -> bytes:
    """Create a Basic Batch ZIP package without embedding raw inputs."""

    output = BytesIO()
    summary_bytes = _batch_summary_workbook(
        summary_rows,
        generated_at=generated_at,
        mapping_preset_name=mapping_preset_name,
    )
    with ZipFile(output, "w", compression=ZIP_DEFLATED) as archive:
        archive.writestr("batch_summary.xlsx", summary_bytes)
        for artifact in file_artifacts:
            folder = safe_batch_name(artifact.folder_name, "file")
            archive.writestr(f"{folder}/report.xlsx", bytes(artifact.workbook_bytes))
            archive.writestr(f"{folder}/export_summary.txt", artifact.export_summary_text.encode("utf-8"))
            archive.writestr(f"{folder}/session.tmcproj.json", bytes(artifact.session_bytes))
            archive.writestr(f"{folder}/mapping_preset.mapping.json", bytes(artifact.mapping_preset_bytes))
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
    confirmed_peak_periods: dict[str, tuple[str, str]] | None = None,
    suggested_am_peak: str = "",
    suggested_pm_peak: str = "",
) -> tuple[BatchSummaryRow, _BatchFileArtifacts]:
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
            setup=setup,
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
        **setup,
        "peak_selection_source": "user_confirmed_batch",
    }
    if "AM" in confirmed_periods:
        confirmed_setup["am_peak_start"], confirmed_setup["am_peak_end"] = confirmed_periods["AM"]
    if "PM" in confirmed_periods:
        confirmed_setup["pm_peak_start"], confirmed_setup["pm_peak_end"] = confirmed_periods["PM"]

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
        use_excel_com_native_charts=False,
        export_mode=export_mode,
        source_file_name=item.file_name,
        generated_at=generated_at,
    )
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
            "use_excel_com_native_charts": False,
            "template_version": TEMPLATE_VERSION,
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
        workbook_filename="report.xlsx",
        pce_factors=result.pce_factors,
        export_settings={"template_version": TEMPLATE_VERSION},
        generated_at=generated_at,
    )
    counts = _qc_counts(result.qc)
    row = BatchSummaryRow(
        file_name=Path(item.file_name).name,
        folder_name=folder_name,
        status="success",
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
        export_file=f"{folder_name}/report.xlsx",
        notes="Auto/suggested peaks confirmed by Batch v1.",
    )
    artifact = _BatchFileArtifacts(
        folder_name=folder_name,
        workbook_bytes=result.workbook_bytes,
        export_summary_text=summary_text,
        session_bytes=session_bytes,
        mapping_preset_bytes=mapping_preset_bytes,
        chart_pngs=chart_pngs,
        diagram_png=diagram_png,
    )
    return row, artifact


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
    for index, item in enumerate(items, start=1):
        safe_stem = safe_batch_name(item.file_name, f"file_{index:02d}")
        folder_name = f"file_{index:02d}_{safe_stem}"
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
                setup=setup,
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
                    folder_name=folder_name,
                    status="success",
                    workbook_bytes=item.workbook_bytes,
                    mapping=active_mapping,
                    suggested_AM_peak=suggested_am,
                    suggested_PM_peak=suggested_pm,
                    confirmed_AM_peak=suggested_am,
                    confirmed_PM_peak=suggested_pm,
                    hourly_period_options=options,
                    total_vehicles=float(result.normalized["count"].sum()) if "count" in result.normalized else 0.0,
                    total_PCU=float(result.normalized["pcu"].sum()) if "pcu" in result.normalized else 0.0,
                    QC_errors=counts["error"],
                    QC_warnings=counts["warning"],
                    QC_info=counts["info"],
                    notes="Suggested peaks are ready for review.",
                )
            )
        except Exception as exc:
            analysis_items.append(
                BatchAnalysisItem(
                    file_name=Path(item.file_name).name,
                    folder_name=folder_name,
                    status="failed",
                    notes=str(exc),
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
) -> BatchResult:
    """Generate the final Batch ZIP using reviewed per-file peak selections."""

    rows: list[BatchSummaryRow] = []
    artifacts: list[_BatchFileArtifacts] = []
    setup = dict(setup or {})
    for item in analysis.items:
        if item.status != "success":
            rows.append(
                BatchSummaryRow(
                    file_name=Path(item.file_name).name,
                    folder_name=item.folder_name,
                    status="failed",
                    notes=item.notes,
                )
            )
            continue

        confirmed_periods = _confirmed_periods_from_labels(item.confirmed_AM_peak, item.confirmed_PM_peak)
        if "AM" not in confirmed_periods or "PM" not in confirmed_periods:
            rows.append(
                BatchSummaryRow(
                    file_name=Path(item.file_name).name,
                    folder_name=item.folder_name,
                    status="failed",
                    suggested_AM_peak=item.suggested_AM_peak,
                    suggested_PM_peak=item.suggested_PM_peak,
                    notes="Confirmed AM/PM peak is missing.",
                )
            )
            continue

        try:
            row, artifact = _process_one_file(
                BatchItem(file_name=item.file_name, workbook_bytes=item.workbook_bytes),
                folder_name=item.folder_name,
                mapping=item.mapping,
                setup=setup,
                pce_factors=pce_factors,
                peak_mode=peak_mode,
                peak_windows=peak_windows,
                export_mode=export_mode,
                generated_at=analysis.generated_at,
                use_template_report_layout=use_template_report_layout,
                confirmed_peak_periods=confirmed_periods,
                suggested_am_peak=item.suggested_AM_peak,
                suggested_pm_peak=item.suggested_PM_peak,
            )
            rows.append(row)
            artifacts.append(artifact)
        except Exception as exc:
            rows.append(
                BatchSummaryRow(
                    file_name=Path(item.file_name).name,
                    folder_name=item.folder_name,
                    status="failed",
                    suggested_AM_peak=item.suggested_AM_peak,
                    suggested_PM_peak=item.suggested_PM_peak,
                    confirmed_AM_peak=item.confirmed_AM_peak,
                    confirmed_PM_peak=item.confirmed_PM_peak,
                    notes=str(exc),
                )
            )

    package = create_batch_package_zip(
        summary_rows=rows,
        file_artifacts=artifacts,
        generated_at=analysis.generated_at,
        mapping_preset_name=analysis.mapping_preset_name,
    )
    return BatchResult(summary_rows=rows, package_bytes=package, generated_at=analysis.generated_at)


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
    )
