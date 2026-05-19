"""Quality-control checks for TMC processing."""

from __future__ import annotations

import pandas as pd

from .constants import AM_WINDOW, DEFAULT_PCE_FACTORS, PM_WINDOW, VEHICLE_CLASSES
from .mapping import clean_mapping
from .pcu import PceFactorIssue, normalize_pce_factors
from .time_utils import minutes_to_time, time_to_minutes


QC_COLUMNS = [
    "severity",
    "category",
    "check",
    "message",
    "detail",
    "affected_field",
    "raw_sheet",
    "movement_code",
]

HIGH_MOVEMENT_SHARE_INFO = 0.45
HIGH_MOVEMENT_SHARE_WARNING = 0.60
HEAVY_VEHICLE_CLASSES = {"LB", "MB", "HB", "LT", "MT", "HT", "TR", "STR"}
HEAVY_VEHICLE_SHARE_INFO = 0.25
HEAVY_VEHICLE_SHARE_WARNING = 0.40
LOW_PHF_INFO = 0.85
LOW_PHF_WARNING = 0.75
PEAK_SPIKE_INFO_RATIO = 1.5
PEAK_SPIKE_WARNING_RATIO = 2.0


def _issue(
    check: str,
    severity: str,
    message: str,
    raw_sheet: str = "",
    movement_code: str = "",
    *,
    category: str = "data_quality",
    detail: str = "",
    affected_field: str = "",
) -> dict[str, str]:
    return {
        "severity": severity,
        "category": category,
        "check": check,
        "message": message,
        "detail": detail,
        "affected_field": affected_field,
        "raw_sheet": raw_sheet,
        "movement_code": movement_code,
    }


def _empty_qc_frame(issues: list[dict[str, str]] | None = None) -> pd.DataFrame:
    return pd.DataFrame(issues or [], columns=QC_COLUMNS)


def _severity_for_share(value: float, *, info_threshold: float, warning_threshold: float) -> str | None:
    if value > warning_threshold:
        return "warning"
    if value > info_threshold:
        return "info"
    return None


def _interval_pcu(normalized: pd.DataFrame) -> pd.DataFrame:
    if normalized.empty:
        return pd.DataFrame(columns=["minute", "pcu"])
    df = normalized[normalized["include_in_peak"]].copy() if "include_in_peak" in normalized else normalized.copy()
    df["minute"] = df["time_start"].map(time_to_minutes)
    df = df.dropna(subset=["minute"])
    if df.empty:
        return pd.DataFrame(columns=["minute", "pcu"])
    return df.groupby("minute", as_index=False).agg(pcu=("pcu", "sum")).sort_values("minute")


def _hourly_pcu_at(interval: pd.DataFrame, start_minute: int) -> float | None:
    hour = interval[(interval["minute"] >= start_minute) & (interval["minute"] < start_minute + 60)]
    if len(hour) < 4:
        return None
    return float(hour["pcu"].sum())


def _time_range_text(start_minute: int) -> str:
    return f"{minutes_to_time(start_minute).strftime('%H:%M')}-{minutes_to_time(start_minute + 60).strftime('%H:%M')}"


def run_qc(
    detected_sheets: list[str],
    mapping: pd.DataFrame,
    normalized: pd.DataFrame,
    peaks: pd.DataFrame | None = None,
    peak_windows: dict[str, tuple[str, str]] | None = None,
    pce_factors: dict[str, float] | None = None,
    raw_sheets: dict[str, pd.DataFrame] | None = None,
    pce_factor_issues: tuple[PceFactorIssue, ...] | list[PceFactorIssue] | None = None,
) -> pd.DataFrame:
    pce_factors = normalize_pce_factors(pce_factors)
    peak_windows = peak_windows or {"AM": AM_WINDOW, "PM": PM_WINDOW}
    original_mapping = mapping.copy()
    mapping = clean_mapping(mapping)
    issues: list[dict[str, str]] = []

    for issue in pce_factor_issues or ():
        issues.append(
            _issue(
                issue.check,
                issue.severity,
                issue.message,
                movement_code=issue.vehicle_class,
                category="pce",
                affected_field="pce_factor",
            )
        )

    pce_overrides = []
    for vehicle_class in VEHICLE_CLASSES:
        selected = pce_factors.get(vehicle_class, DEFAULT_PCE_FACTORS.get(vehicle_class))
        default = DEFAULT_PCE_FACTORS.get(vehicle_class)
        if selected is not None and default is not None and abs(float(selected) - float(default)) > 1e-12:
            pce_overrides.append(f"{vehicle_class}: default {default:g}, selected {float(selected):g}")
    if pce_overrides:
        issues.append(
            _issue(
                "pce_override",
                "info",
                "Custom PCE factors are used.",
                category="pce_override",
                detail="; ".join(pce_overrides),
                affected_field="pce_factor",
            )
        )

    included_mapping = mapping[mapping["include_in_report"]]
    mapped_sheets = set(included_mapping.loc[included_mapping["movement_code"].str.strip() != "", "raw_sheet"])
    for sheet in detected_sheets:
        if sheet not in mapped_sheets:
            excluded = mapping[(mapping["raw_sheet"] == sheet) & (~mapping["include_in_report"])]
            if excluded.empty:
                issues.append(_issue("unmapped_raw_sheet", "error", "Detected raw sheet has no movement mapping.", raw_sheet=sheet, category="mapping"))

    duplicate_counts = included_mapping.loc[included_mapping["movement_code"].str.strip() != ""].groupby("movement_code")["raw_sheet"].count()
    for movement_code, count in duplicate_counts.items():
        if int(count) > 1:
            issues.append(
                _issue(
                    "aggregated_output_movement",
                    "info",
                    f"{movement_code} is aggregated from {int(count)} source streams.",
                    movement_code=str(movement_code),
                    category="movement_aggregation",
                    detail=f"source_stream_count={int(count)}",
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
                category="mapping",
                affected_field="include_in_report",
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
                    category="mapping",
                    affected_field="source_stream",
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
                        category="survey_data",
                        affected_field="vehicle_class",
                    )
                )

    if normalized.empty:
        return _empty_qc_frame(issues)

    missing_vehicle = sorted(set(normalized["vehicle_class"].dropna()) - set(VEHICLE_CLASSES))
    for vehicle_class in missing_vehicle:
        issues.append(
            _issue(
                "missing_vehicle_class",
                "warning",
                f"Unexpected vehicle class: {vehicle_class}",
                category="survey_data",
                affected_field="vehicle_class",
            )
        )

    missing_pce = sorted(set(normalized.loc[normalized["pce_factor"].isna(), "vehicle_class"].dropna()))
    missing_pce.extend(sorted(set(normalized["vehicle_class"].dropna()) - set(pce_factors)))
    missing_pce = sorted(set(missing_pce))
    for vehicle_class in missing_pce:
        issues.append(
            _issue(
                "missing_pce_factor",
                "warning",
                f"Missing PCE factor for vehicle class: {vehicle_class}",
                category="pce",
                movement_code=vehicle_class,
                affected_field="pce_factor",
            )
        )

    report_rows = normalized[normalized["include_in_report"]].copy() if "include_in_report" in normalized else normalized.copy()
    movement_totals = report_rows.groupby("movement_code", dropna=False, as_index=False).agg(count=("count", "sum"), pcu=("pcu", "sum"))
    for _, row in movement_totals.iterrows():
        movement_code = str(row["movement_code"])
        count_total = float(row["count"] or 0)
        pcu_total = float(row["pcu"] or 0)
        if movement_code.strip() and (count_total == 0 or pcu_total == 0):
            issues.append(
                _issue(
                    "zero_volume_mapped_movement",
                    "warning",
                    "Movement appears in report but has zero volume.",
                    movement_code=movement_code,
                    category="movement_volume",
                    detail=f"total_vehicle_count={count_total:g}; total_pcu={pcu_total:g}",
                )
            )

    total_pcu = float(report_rows["pcu"].sum()) if "pcu" in report_rows else 0.0
    if total_pcu > 0:
        for _, row in movement_totals.iterrows():
            movement_code = str(row["movement_code"])
            if not movement_code.strip():
                continue
            movement_pcu = float(row["pcu"] or 0)
            share = movement_pcu / total_pcu
            severity = _severity_for_share(
                share,
                info_threshold=HIGH_MOVEMENT_SHARE_INFO,
                warning_threshold=HIGH_MOVEMENT_SHARE_WARNING,
            )
            if severity:
                issues.append(
                    _issue(
                        "high_single_movement_share",
                        severity,
                        "One movement accounts for an unusually high share of total PCU.",
                        movement_code=movement_code,
                        category="movement_volume",
                        detail=f"movement_pcu={movement_pcu:g}; total_pcu={total_pcu:g}; share={share:.1%}",
                    )
                )

    if {"vehicle_class", "count"}.issubset(report_rows.columns):
        total_count = float(report_rows["count"].sum())
        heavy_count = float(report_rows.loc[report_rows["vehicle_class"].isin(HEAVY_VEHICLE_CLASSES), "count"].sum())
        if total_count > 0:
            heavy_share = heavy_count / total_count
            severity = _severity_for_share(
                heavy_share,
                info_threshold=HEAVY_VEHICLE_SHARE_INFO,
                warning_threshold=HEAVY_VEHICLE_SHARE_WARNING,
            )
            if severity:
                issues.append(
                    _issue(
                        "high_heavy_vehicle_share",
                        severity,
                        "Heavy vehicle share is unusually high.",
                        category="vehicle_composition",
                        detail=f"heavy_vehicle_count={heavy_count:g}; total_vehicle_count={total_count:g}; share={heavy_share:.1%}",
                        affected_field="vehicle_class",
                    )
                )

    if peaks is not None and not peaks.empty:
        survey_minutes = normalized["time_start"].map(time_to_minutes).dropna()
        min_minute = int(survey_minutes.min()) if not survey_minutes.empty else None
        max_minute = int(survey_minutes.max()) + 15 if not survey_minutes.empty else None
        interval_pcu = _interval_pcu(normalized)
        for _, row in peaks.iterrows():
            start = time_to_minutes(row["peak_start"])
            end = time_to_minutes(row["peak_end"])
            period = str(row["period"])
            if min_minute is not None and (start is None or end is None or start < min_minute or end > max_minute):
                issues.append(_issue("peak_outside_survey_time", "error", f"{period} peak falls outside survey time.", category="peak"))
            window = peak_windows.get(period.upper())
            if window:
                window_start = time_to_minutes(window[0])
                window_end = time_to_minutes(window[1])
                if start is not None and end is not None and (start < window_start or end > window_end):
                    issues.append(_issue("peak_outside_survey_time", "error", f"{period} peak falls outside configured peak window.", category="peak"))

            phf = pd.to_numeric(pd.Series([row.get("phf")]), errors="coerce").iloc[0] if "phf" in row else None
            if pd.notna(phf):
                severity = "warning" if float(phf) < LOW_PHF_WARNING else ("info" if float(phf) < LOW_PHF_INFO else None)
                if severity:
                    issues.append(
                        _issue(
                            "low_peak_phf",
                            severity,
                            f"{period} peak PHF is unusually low.",
                            category="peak_phf",
                            detail=f"phf={float(phf):.3f}; info_threshold<{LOW_PHF_INFO:g}; warning_threshold<{LOW_PHF_WARNING:g}",
                            affected_field="phf",
                        )
                    )

            if start is not None and end is not None and not interval_pcu.empty:
                peak_pcu = float(row.get("hourly_pcu") or 0)
                adjacent_values = []
                previous = _hourly_pcu_at(interval_pcu, start - 60)
                following = _hourly_pcu_at(interval_pcu, end)
                if previous is not None:
                    adjacent_values.append(previous)
                if following is not None:
                    adjacent_values.append(following)
                adjacent_average = sum(adjacent_values) / len(adjacent_values) if adjacent_values else 0
                if adjacent_average > 0:
                    ratio = peak_pcu / adjacent_average
                    severity = _severity_for_share(
                        ratio,
                        info_threshold=PEAK_SPIKE_INFO_RATIO,
                        warning_threshold=PEAK_SPIKE_WARNING_RATIO,
                    )
                    if severity:
                        issues.append(
                            _issue(
                                "peak_spike_vs_adjacent_hours",
                                severity,
                                f"{period} peak hour is much higher than adjacent hourly volumes.",
                                category="peak_spike",
                                detail=(
                                    f"peak={_time_range_text(start)}; peak_pcu={peak_pcu:g}; "
                                    f"adjacent_average_pcu={adjacent_average:g}; ratio={ratio:.2f}"
                                ),
                                affected_field="hourly_pcu",
                            )
                        )

    return _empty_qc_frame(issues)
