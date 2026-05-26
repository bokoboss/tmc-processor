"""High-level processing orchestration."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

import pandas as pd

from .exporter import export_workbook
from .mapping import clean_mapping, mapping_processing_block_reason, validate_mapping_for_processing, validate_mapping_scheme
from .movement_scheme import MOVEMENT_SCHEME_V1, MOVEMENT_SCHEME_V2
from .metadata import setup_with_metadata
from .normalizer import normalize
from .pcu import get_default_pce_factors, validate_pce_factors
from .peaks import (
    PEAK_SELECTION_AUTO,
    PEAK_SELECTION_USER_CONFIRMED,
    confirmed_peak_periods_from_setup,
    confirmed_peak_phf,
    detect_peak_phf,
)
from .constants import AM_WINDOW, DEFAULT_PEAK_MODE, PM_WINDOW
from .qc import run_qc
from .summaries import hourly_movement_pcu, hourly_summary, movement_summary, vehicle_composition


@dataclass
class ProcessingResult:
    normalized: pd.DataFrame
    qc: pd.DataFrame
    hourly: pd.DataFrame
    movement: pd.DataFrame
    vehicle: pd.DataFrame
    peaks: pd.DataFrame
    workbook_bytes: bytes
    pce_factors: dict[str, float] = field(default_factory=get_default_pce_factors)


@dataclass
class V2DryRunResult:
    normalized: pd.DataFrame
    qc: pd.DataFrame
    hourly: pd.DataFrame
    hourly_movement_pcu: pd.DataFrame
    movement: pd.DataFrame
    vehicle: pd.DataFrame
    peaks: pd.DataFrame
    workbook_bytes: bytes = b""
    movement_code_scheme: str = MOVEMENT_SCHEME_V2
    pce_factors: dict[str, float] = field(default_factory=get_default_pce_factors)


def _raise_mapping_scheme_issues(mapping: pd.DataFrame, movement_code_scheme: str) -> None:
    scheme_issues = validate_mapping_scheme(mapping, movement_code_scheme)
    if scheme_issues:
        raise ValueError("; ".join(scheme_issues))


def _v2_mapping_dry_run_issues(detected_sheets: list[str], mapping: pd.DataFrame) -> pd.DataFrame:
    cleaned = clean_mapping(mapping)
    issues = []
    for sheet in detected_sheets:
        sheet_rows = cleaned[cleaned["raw_sheet"] == sheet]
        included_rows = sheet_rows[sheet_rows["include_in_report"]] if not sheet_rows.empty else sheet_rows
        if sheet_rows.empty or included_rows.empty:
            issues.append(
                {
                    "raw_sheet": sheet,
                    "field": "movement_code",
                    "message": "Detected raw sheet requires an approach_movement mapping before v2 dry-run.",
                }
            )
            continue
        for _, row in included_rows.iterrows():
            if not str(row.get("movement_code") or "").strip():
                issues.append(
                    {
                        "raw_sheet": sheet,
                        "field": "movement_code",
                        "message": "Detected raw sheet requires movement_code before v2 dry-run.",
                    }
                )
    return pd.DataFrame(issues, columns=["raw_sheet", "field", "message"])


def process_tmc(
    raw_sheets: dict[str, pd.DataFrame],
    mapping: pd.DataFrame,
    setup: dict[str, Any],
    detected_sheets: list[str] | None = None,
    peak_mode: str = DEFAULT_PEAK_MODE,
    peak_windows: dict[str, tuple[str, str]] | None = None,
    confirmed_peak_periods: dict[str, tuple[str, str]] | None = None,
    generate_workbook: bool = True,
    use_template_report_layout: bool = False,
    use_excel_com_native_charts: bool = False,
    pce_factors: dict[str, float] | None = None,
    export_mode: str | None = None,
    source_file_name: str | None = None,
    generated_at: str | None = None,
) -> ProcessingResult:
    detected_sheets = detected_sheets or list(raw_sheets)
    movement_code_scheme = str(setup.get("movement_code_scheme") or MOVEMENT_SCHEME_V1)
    block_reason = mapping_processing_block_reason(movement_code_scheme)
    if block_reason:
        raise ValueError(block_reason)
    _raise_mapping_scheme_issues(mapping, movement_code_scheme)
    mapping_issues = validate_mapping_for_processing(detected_sheets, mapping)
    if not mapping_issues.empty:
        issue_text = "; ".join(
            f"{row.raw_sheet}: {row.field}" for row in mapping_issues.itertuples(index=False)
        )
        raise ValueError(f"Mapping is incomplete for detected raw sheets: {issue_text}")
    setup = setup_with_metadata(setup)
    pce_validation = validate_pce_factors(pce_factors)
    selected_pce_factors = pce_validation.factors
    normalized = normalize(raw_sheets=raw_sheets, mapping=mapping, setup=setup, pce_factors=selected_pce_factors)
    peak_windows = peak_windows or {"AM": AM_WINDOW, "PM": PM_WINDOW}
    setup = {
        **setup,
        "peak_mode": peak_mode,
        "am_peak_window_start": peak_windows["AM"][0],
        "am_peak_window_end": peak_windows["AM"][1],
        "pm_peak_window_start": peak_windows["PM"][0],
        "pm_peak_window_end": peak_windows["PM"][1],
    }
    stored_peak_periods = confirmed_peak_periods_from_setup(setup)
    if (
        not confirmed_peak_periods
        and str(setup.get("peak_selection_source") or "").casefold() == PEAK_SELECTION_USER_CONFIRMED
        and stored_peak_periods
    ):
        confirmed_peak_periods = stored_peak_periods
    if confirmed_peak_periods:
        selection_source = setup.get("peak_selection_source") or PEAK_SELECTION_USER_CONFIRMED
        setup = {
            **setup,
            "am_peak_start": confirmed_peak_periods.get("AM", ("", ""))[0],
            "am_peak_end": confirmed_peak_periods.get("AM", ("", ""))[1],
            "pm_peak_start": confirmed_peak_periods.get("PM", ("", ""))[0],
            "pm_peak_end": confirmed_peak_periods.get("PM", ("", ""))[1],
            "peak_selection_source": selection_source,
        }
        peaks = confirmed_peak_phf(
            normalized,
            peak_periods=confirmed_peak_periods,
            peak_mode=peak_mode,
            peak_selection_source=selection_source,
        )
    else:
        setup = {**setup, "peak_selection_source": setup.get("peak_selection_source") or PEAK_SELECTION_AUTO}
        peaks = detect_peak_phf(normalized, windows=peak_windows, peak_mode=peak_mode)
    qc = run_qc(
        detected_sheets=detected_sheets,
        mapping=mapping,
        normalized=normalized,
        peaks=peaks,
        peak_windows=peak_windows,
        pce_factors=selected_pce_factors,
        pce_factor_issues=pce_validation.issues,
        raw_sheets=raw_sheets,
    )
    hourly = hourly_summary(normalized)
    movement = movement_summary(normalized)
    vehicle = vehicle_composition(normalized)
    workbook_bytes = b""
    if generate_workbook:
        workbook_bytes = export_workbook(
            setup,
            mapping,
            normalized,
            qc,
            hourly,
            movement,
            vehicle,
            peaks,
            pce_factors=selected_pce_factors,
            use_template_report_layout=use_template_report_layout,
            use_excel_com_native_charts=use_excel_com_native_charts,
            export_mode=export_mode,
            source_file_name=source_file_name,
            generated_at=generated_at,
        )
    return ProcessingResult(normalized, qc, hourly, movement, vehicle, peaks, workbook_bytes, selected_pce_factors)


def process_tmc_dry_run_v2(
    raw_sheets: dict[str, pd.DataFrame],
    mapping: pd.DataFrame,
    setup: dict[str, Any],
    detected_sheets: list[str] | None = None,
    peak_mode: str = DEFAULT_PEAK_MODE,
    peak_windows: dict[str, tuple[str, str]] | None = None,
    confirmed_peak_periods: dict[str, tuple[str, str]] | None = None,
    pce_factors: dict[str, float] | None = None,
) -> V2DryRunResult:
    """Run the approach_movement v2 path for normalization and summaries only."""

    detected_sheets = detected_sheets or list(raw_sheets)
    setup = {**setup, "movement_code_scheme": MOVEMENT_SCHEME_V2}
    _raise_mapping_scheme_issues(mapping, MOVEMENT_SCHEME_V2)
    mapping_issues = _v2_mapping_dry_run_issues(detected_sheets, mapping)
    if not mapping_issues.empty:
        issue_text = "; ".join(
            f"{row.raw_sheet}: {row.field}" for row in mapping_issues.itertuples(index=False)
        )
        raise ValueError(f"Mapping is incomplete for detected raw sheets: {issue_text}")

    setup = setup_with_metadata(setup)
    pce_validation = validate_pce_factors(pce_factors)
    selected_pce_factors = pce_validation.factors
    normalized = normalize(raw_sheets=raw_sheets, mapping=mapping, setup=setup, pce_factors=selected_pce_factors)
    peak_windows = peak_windows or {"AM": AM_WINDOW, "PM": PM_WINDOW}
    setup = {
        **setup,
        "peak_mode": peak_mode,
        "am_peak_window_start": peak_windows["AM"][0],
        "am_peak_window_end": peak_windows["AM"][1],
        "pm_peak_window_start": peak_windows["PM"][0],
        "pm_peak_window_end": peak_windows["PM"][1],
    }
    stored_peak_periods = confirmed_peak_periods_from_setup(setup)
    if (
        not confirmed_peak_periods
        and str(setup.get("peak_selection_source") or "").casefold() == PEAK_SELECTION_USER_CONFIRMED
        and stored_peak_periods
    ):
        confirmed_peak_periods = stored_peak_periods
    if confirmed_peak_periods:
        selection_source = setup.get("peak_selection_source") or PEAK_SELECTION_USER_CONFIRMED
        peaks = confirmed_peak_phf(
            normalized,
            peak_periods=confirmed_peak_periods,
            peak_mode=peak_mode,
            peak_selection_source=selection_source,
        )
    else:
        peaks = detect_peak_phf(normalized, windows=peak_windows, peak_mode=peak_mode)
    qc = run_qc(
        detected_sheets=detected_sheets,
        mapping=mapping,
        normalized=normalized,
        peaks=peaks,
        peak_windows=peak_windows,
        pce_factors=selected_pce_factors,
        pce_factor_issues=pce_validation.issues,
        raw_sheets=raw_sheets,
    )
    hourly = hourly_summary(normalized)
    hourly_movement = hourly_movement_pcu(normalized, mapping, movement_code_scheme=MOVEMENT_SCHEME_V2)
    movement = movement_summary(normalized)
    vehicle = vehicle_composition(normalized)
    return V2DryRunResult(
        normalized=normalized,
        qc=qc,
        hourly=hourly,
        hourly_movement_pcu=hourly_movement,
        movement=movement,
        vehicle=vehicle,
        peaks=peaks,
        workbook_bytes=b"",
        movement_code_scheme=MOVEMENT_SCHEME_V2,
        pce_factors=selected_pce_factors,
    )
