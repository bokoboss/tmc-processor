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
from .peaks import PEAK_SELECTION_USER_CONFIRMED
from .pipeline import ProcessingResult, process_tmc
from .session import build_project_session, session_to_json_bytes
from .summaries import hourly_movement_pcu, vehicle_composition_report


BATCH_PACKAGE_MIME = "application/zip"
SAFE_BATCH_EXPORT_MODE = "Safe PNG Export Mode - Batch v1"


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
    summary = pd.DataFrame([row.__dict__ for row in rows])
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
) -> tuple[BatchSummaryRow, _BatchFileArtifacts]:
    raw_sheets = load_detected_sheets(BytesIO(item.workbook_bytes))
    detected_sheets = list(raw_sheets)
    apply_result = apply_mapping_preset_to_detected_sheets(
        build_mapping_preset(mapping, preset_name="Batch Mapping Preset"),
        detected_sheets,
    )
    active_mapping = apply_result.mapping

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
    confirmed_setup = {
        **setup,
        "peak_selection_source": PEAK_SELECTION_USER_CONFIRMED,
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
    setup = dict(setup or {})
    source_mapping = clean_mapping(mapping if mapping is not None else pd.DataFrame())
    if source_mapping.empty and mapping_preset:
        source_mapping = apply_mapping_preset_to_detected_sheets(mapping_preset, []).mapping

    rows: list[BatchSummaryRow] = []
    artifacts: list[_BatchFileArtifacts] = []
    for index, item in enumerate(items, start=1):
        safe_stem = safe_batch_name(item.file_name, f"file_{index:02d}")
        folder_name = f"file_{index:02d}_{safe_stem}"
        try:
            if source_mapping.empty and mapping_preset:
                detected = load_detected_sheets(BytesIO(item.workbook_bytes))
                active_mapping = apply_mapping_preset_to_detected_sheets(mapping_preset, list(detected)).mapping
            else:
                active_mapping = source_mapping
            row, artifact = _process_one_file(
                item,
                folder_name=folder_name,
                mapping=active_mapping,
                setup=setup,
                pce_factors=pce_factors,
                peak_mode=peak_mode,
                peak_windows=peak_windows,
                export_mode=export_mode,
                generated_at=generated_at,
                use_template_report_layout=use_template_report_layout,
            )
            rows.append(row)
            artifacts.append(artifact)
        except Exception as exc:
            rows.append(
                BatchSummaryRow(
                    file_name=Path(item.file_name).name,
                    folder_name=folder_name,
                    status="failed",
                    notes=str(exc),
                )
            )

    package = create_batch_package_zip(
        summary_rows=rows,
        file_artifacts=artifacts,
        generated_at=generated_at,
        mapping_preset_name=mapping_preset_name,
    )
    return BatchResult(summary_rows=rows, package_bytes=package, generated_at=generated_at)
