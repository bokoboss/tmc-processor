"""Quality-control checks for TMC processing."""

from __future__ import annotations

import pandas as pd

from .constants import AM_WINDOW, PM_WINDOW, PCE_FACTORS, VEHICLE_CLASSES
from .mapping import clean_mapping
from .time_utils import time_to_minutes


def _issue(check: str, severity: str, message: str, raw_sheet: str = "", movement_code: str = "") -> dict[str, str]:
    return {
        "check": check,
        "severity": severity,
        "raw_sheet": raw_sheet,
        "movement_code": movement_code,
        "message": message,
    }


def run_qc(
    detected_sheets: list[str],
    mapping: pd.DataFrame,
    normalized: pd.DataFrame,
    peaks: pd.DataFrame | None = None,
    peak_windows: dict[str, tuple[str, str]] | None = None,
    pce_factors: dict[str, float] | None = None,
    raw_sheets: dict[str, pd.DataFrame] | None = None,
) -> pd.DataFrame:
    pce_factors = pce_factors or PCE_FACTORS
    peak_windows = peak_windows or {"AM": AM_WINDOW, "PM": PM_WINDOW}
    original_mapping = mapping.copy()
    mapping = clean_mapping(mapping)
    issues: list[dict[str, str]] = []

    included_mapping = mapping[mapping["include_in_report"]]
    mapped_sheets = set(included_mapping.loc[included_mapping["movement_code"].str.strip() != "", "raw_sheet"])
    for sheet in detected_sheets:
        if sheet not in mapped_sheets:
            excluded = mapping[(mapping["raw_sheet"] == sheet) & (~mapping["include_in_report"])]
            if excluded.empty:
                issues.append(_issue("unmapped_raw_sheet", "error", "Detected raw sheet has no movement mapping.", raw_sheet=sheet))

    duplicate_counts = included_mapping.loc[included_mapping["movement_code"].str.strip() != ""].groupby("movement_code")["raw_sheet"].count()
    for movement_code, count in duplicate_counts.items():
        if int(count) > 1:
            issues.append(
                _issue(
                    "aggregated_output_movement",
                    "info",
                    f"{movement_code} is aggregated from {int(count)} source streams.",
                    movement_code=str(movement_code),
                )
            )

    excluded_rows = mapping[~mapping["include_in_report"]]
    for _, row in excluded_rows.iterrows():
        issues.append(
            _issue(
                "source_row_excluded_from_report",
                "info",
                "Mapping row is excluded from report outputs.",
                raw_sheet=str(row["raw_sheet"]),
                movement_code=str(row["movement_code"]),
            )
        )

    if "source_stream" in original_mapping.columns:
        blank_stream_rows = original_mapping[original_mapping["source_stream"].fillna("").astype(str).str.strip() == ""]
        for _, row in blank_stream_rows.iterrows():
            issues.append(
                _issue(
                    "blank_source_stream_defaulted",
                    "info",
                    "Blank source_stream defaulted to mainline.",
                    raw_sheet=str(row.get("raw_sheet", "")),
                    movement_code=str(row.get("movement_code", row.get("output_movement_code", ""))),
                )
            )

    if raw_sheets:
        for sheet, raw in raw_sheets.items():
            if "vehicle_class" in raw.columns:
                detected_classes = set(raw["vehicle_class"].dropna())
            else:
                detected_classes = {str(column).strip() for column in raw.columns} & set(VEHICLE_CLASSES)
            if not detected_classes:
                issues.append(
                    _issue(
                        "missing_vehicle_class",
                        "error",
                        "Raw sheet has no recognized vehicle-class columns.",
                        raw_sheet=sheet,
                    )
                )

    if normalized.empty:
        return pd.DataFrame(issues, columns=["check", "severity", "raw_sheet", "movement_code", "message"])

    missing_vehicle = sorted(set(normalized["vehicle_class"].dropna()) - set(VEHICLE_CLASSES))
    for vehicle_class in missing_vehicle:
        issues.append(_issue("missing_vehicle_class", "error", f"Unexpected vehicle class: {vehicle_class}"))

    missing_pce = sorted(set(normalized.loc[normalized["pce_factor"].isna(), "vehicle_class"].dropna()))
    missing_pce.extend(sorted(set(normalized["vehicle_class"].dropna()) - set(pce_factors)))
    missing_pce = sorted(set(missing_pce))
    for vehicle_class in missing_pce:
        issues.append(_issue("missing_pce_factor", "error", f"Missing PCE factor for vehicle class: {vehicle_class}"))

    movement_totals = normalized.groupby(["raw_sheet", "movement_code"], dropna=False, as_index=False)["count"].sum()
    for _, row in movement_totals.iterrows():
        movement_code = str(row["movement_code"])
        if movement_code.strip() and row["count"] == 0:
            issues.append(
                _issue(
                    "mapped_movement_zero_total",
                    "warning",
                    "Mapped movement has zero total volume.",
                    raw_sheet=str(row["raw_sheet"]),
                    movement_code=movement_code,
                )
            )

    if peaks is not None and not peaks.empty:
        survey_minutes = normalized["time_start"].map(time_to_minutes).dropna()
        min_minute = int(survey_minutes.min()) if not survey_minutes.empty else None
        max_minute = int(survey_minutes.max()) + 15 if not survey_minutes.empty else None
        for _, row in peaks.iterrows():
            start = time_to_minutes(row["peak_start"])
            end = time_to_minutes(row["peak_end"])
            period = str(row["period"])
            if min_minute is not None and (start is None or end is None or start < min_minute or end > max_minute):
                issues.append(_issue("peak_outside_survey_time", "error", f"{period} peak falls outside survey time."))
            window = peak_windows.get(period.upper())
            if window:
                window_start = time_to_minutes(window[0])
                window_end = time_to_minutes(window[1])
                if start is not None and end is not None and (start < window_start or end > window_end):
                    issues.append(_issue("peak_outside_survey_time", "error", f"{period} peak falls outside configured peak window."))

    return pd.DataFrame(issues, columns=["check", "severity", "raw_sheet", "movement_code", "message"])
