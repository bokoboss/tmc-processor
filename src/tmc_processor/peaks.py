"""AM/PM peak hour and PHF calculations."""

from __future__ import annotations

import pandas as pd

from .constants import (
    AM_WINDOW,
    DEFAULT_PEAK_MODE,
    PEAK_INTERVALS,
    PEAK_MODE_FIXED_HOURLY,
    PEAK_MODE_OPTIONS,
    PEAK_MODE_ROLLING_60MIN,
    PM_WINDOW,
)
from .time_utils import minutes_to_time, time_to_minutes


PEAK_SELECTION_AUTO = "auto_suggested"
PEAK_SELECTION_USER_CONFIRMED = "user_confirmed"
PEAK_SELECTION_TEMPLATE_DEFAULT = "template_default"

PEAK_SETUP_KEYS = {
    "AM": ("am_peak_start", "am_peak_end"),
    "PM": ("pm_peak_start", "pm_peak_end"),
}


def _window_minutes(window: tuple[str, str]) -> tuple[int, int]:
    start = time_to_minutes(window[0])
    end = time_to_minutes(window[1])
    if start is None or end is None:
        raise ValueError(f"Invalid peak window: {window}")
    return start, end


def _present(value) -> bool:
    if value is None:
        return False
    if isinstance(value, float) and pd.isna(value):
        return False
    return str(value).strip() != ""


def confirmed_peak_periods_from_setup(setup: dict) -> dict[str, tuple[str, str]]:
    """Extract explicitly stored AM/PM peak periods from setup/export state."""

    periods = {}
    for period, (start_key, end_key) in PEAK_SETUP_KEYS.items():
        start = setup.get(start_key)
        end = setup.get(end_key)
        if _present(start) and _present(end):
            periods[period] = (start, end)
    return periods


def peak_periods_from_frame(peaks: pd.DataFrame) -> dict[str, tuple[str, str]]:
    periods = {}
    if peaks.empty or not {"period", "peak_start", "peak_end"}.issubset(peaks.columns):
        return periods
    for _, row in peaks.iterrows():
        period = str(row.get("period") or "").upper()
        start = row.get("peak_start")
        end = row.get("peak_end")
        if period in PEAK_SETUP_KEYS and _present(start) and _present(end):
            periods[period] = (start, end)
    return periods


def resolve_effective_peak_periods(
    *,
    selected_peak_periods: dict[str, tuple[str, str]] | None = None,
    recommended_peak_periods: dict[str, tuple[str, str]] | None = None,
    template_default_peak_periods: dict[str, tuple[str, str]] | None = None,
) -> tuple[dict[str, tuple[str, str]], str]:
    for periods, source in (
        (selected_peak_periods, PEAK_SELECTION_USER_CONFIRMED),
        (recommended_peak_periods, PEAK_SELECTION_AUTO),
        (template_default_peak_periods, PEAK_SELECTION_TEMPLATE_DEFAULT),
    ):
        cleaned = {period: value for period, value in (periods or {}).items() if period in PEAK_SETUP_KEYS}
        if cleaned:
            return cleaned, source
    return {}, ""


def _candidate_starts(interval: pd.DataFrame, window_start: int, window_end: int, peak_mode: str) -> list[int]:
    if peak_mode == PEAK_MODE_FIXED_HOURLY:
        first_hour = ((window_start + 59) // 60) * 60
        return list(range(first_hour, window_end - 59, 60))
    if peak_mode == PEAK_MODE_ROLLING_60MIN:
        candidates = interval[(interval["minute"] >= window_start) & (interval["minute"] < window_end)].copy()
        return [int(minute) for minute in candidates["minute"]]
    raise ValueError(f"Invalid peak calculation mode: {peak_mode}")


def detect_peak_phf(
    normalized: pd.DataFrame,
    windows: dict[str, tuple[str, str]] | None = None,
    peak_mode: str = DEFAULT_PEAK_MODE,
) -> pd.DataFrame:
    if peak_mode not in PEAK_MODE_OPTIONS:
        raise ValueError(f"Invalid peak calculation mode: {peak_mode}")
    windows = windows or {"AM": AM_WINDOW, "PM": PM_WINDOW}
    columns = [
        "period",
        "peak_mode",
        "peak_start",
        "peak_end",
        "hourly_pcu",
        "max_15min_pcu",
        "phf",
        "peak_selection_source",
    ]
    if normalized.empty:
        return pd.DataFrame(columns=columns)

    df = normalized[normalized["include_in_peak"]].copy()
    df["minute"] = df["time_start"].map(time_to_minutes)
    df = df.dropna(subset=["minute"])
    interval = df.groupby("minute", as_index=False).agg(pcu=("pcu", "sum"))
    interval = interval.sort_values("minute")

    rows = []
    for period, window in windows.items():
        window_start, window_end = _window_minutes(window)
        best = None
        for minute in _candidate_starts(interval, window_start, window_end, peak_mode):
            hour = interval[(interval["minute"] >= minute) & (interval["minute"] < minute + 60)]
            if len(hour) < PEAK_INTERVALS or minute + 60 > window_end:
                continue
            hourly_pcu = hour["pcu"].sum()
            max_15 = hour["pcu"].max()
            phf = hourly_pcu / (PEAK_INTERVALS * max_15) if max_15 else 0
            candidate = {
                "period": period,
                "peak_mode": peak_mode,
                "peak_start": minutes_to_time(int(minute)),
                "peak_end": minutes_to_time(int(minute) + 60),
                "hourly_pcu": hourly_pcu,
                "max_15min_pcu": max_15,
                "phf": phf,
                "peak_selection_source": PEAK_SELECTION_AUTO,
            }
            if best is None or candidate["hourly_pcu"] > best["hourly_pcu"]:
                best = candidate
        if best is not None:
            rows.append(best)
    return pd.DataFrame(rows, columns=columns)


def confirmed_peak_phf(
    normalized: pd.DataFrame,
    peak_periods: dict[str, tuple[str, str]],
    peak_mode: str = DEFAULT_PEAK_MODE,
    peak_selection_source: str = PEAK_SELECTION_USER_CONFIRMED,
) -> pd.DataFrame:
    """Calculate PHF rows for explicitly selected AM/PM peak periods."""

    if peak_mode not in PEAK_MODE_OPTIONS:
        raise ValueError(f"Invalid peak calculation mode: {peak_mode}")
    columns = [
        "period",
        "peak_mode",
        "peak_start",
        "peak_end",
        "hourly_pcu",
        "max_15min_pcu",
        "phf",
        "peak_selection_source",
    ]
    if normalized.empty:
        return pd.DataFrame(columns=columns)

    df = normalized[normalized["include_in_peak"]].copy()
    df["minute"] = df["time_start"].map(time_to_minutes)
    df = df.dropna(subset=["minute"])
    interval = df.groupby("minute", as_index=False).agg(pcu=("pcu", "sum")).sort_values("minute")

    rows = []
    for period in ["AM", "PM"]:
        if period not in peak_periods:
            continue
        start_minute, end_minute = _window_minutes(peak_periods[period])
        if end_minute <= start_minute:
            raise ValueError(f"Invalid confirmed {period} peak period: {peak_periods[period]}")
        selected = interval[(interval["minute"] >= start_minute) & (interval["minute"] < end_minute)]
        hourly_pcu = float(selected["pcu"].sum()) if not selected.empty else 0.0
        max_15 = float(selected["pcu"].max()) if not selected.empty else 0.0
        rows.append(
            {
                "period": period,
                "peak_mode": peak_mode,
                "peak_start": minutes_to_time(start_minute),
                "peak_end": minutes_to_time(end_minute),
                "hourly_pcu": hourly_pcu,
                "max_15min_pcu": max_15,
                "phf": hourly_pcu / (PEAK_INTERVALS * max_15) if max_15 else 0,
                "peak_selection_source": peak_selection_source,
            }
        )
    return pd.DataFrame(rows, columns=columns)
