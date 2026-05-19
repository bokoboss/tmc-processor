"""Passenger car unit calculations and PCE factor handling."""

from __future__ import annotations

from dataclasses import dataclass
from math import isclose
from typing import Any, Iterable, Mapping

import pandas as pd

from .constants import DEFAULT_PCE_FACTORS, PCE_FACTORS, VEHICLE_CLASSES


@dataclass(frozen=True)
class PceFactorIssue:
    check: str
    severity: str
    vehicle_class: str
    message: str


@dataclass(frozen=True)
class PceValidationResult:
    factors: dict[str, float]
    issues: tuple[PceFactorIssue, ...] = ()

    @property
    def warnings(self) -> tuple[str, ...]:
        return tuple(issue.message for issue in self.issues if issue.severity in {"warning", "error"})


def get_default_pce_factors() -> dict[str, float]:
    """Return a copy of the standard PCE factors used by default."""

    return dict(DEFAULT_PCE_FACTORS)


def _same_factor(left: Any, right: Any) -> bool:
    try:
        return isclose(float(left), float(right), rel_tol=0, abs_tol=1e-12)
    except (TypeError, ValueError):
        return False


def _coerce_factor(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, str) and not value.strip():
        return None
    try:
        parsed = float(value)
    except (TypeError, ValueError):
        return None
    if pd.isna(parsed):
        return None
    return parsed


def _dataframe_factor_items(frame: pd.DataFrame) -> Iterable[tuple[Any, Any]]:
    if frame.empty:
        return []
    columns_by_lower = {str(column).strip().lower(): column for column in frame.columns}
    class_column = columns_by_lower.get("vehicle_class") or columns_by_lower.get("class")
    factor_column = columns_by_lower.get("pce_factor") or columns_by_lower.get("pce")
    if class_column is None or factor_column is None:
        return []
    return [(row[class_column], row[factor_column]) for _, row in frame.iterrows()]


def _factor_items(pce_factors: Any) -> Iterable[tuple[Any, Any]]:
    if pce_factors is None:
        return []
    if isinstance(pce_factors, pd.DataFrame):
        return _dataframe_factor_items(pce_factors)
    if isinstance(pce_factors, Mapping):
        if "vehicle_class" in pce_factors and ("pce_factor" in pce_factors or "pce" in pce_factors):
            return [(pce_factors.get("vehicle_class"), pce_factors.get("pce_factor", pce_factors.get("pce")))]
        return pce_factors.items()
    if isinstance(pce_factors, Iterable) and not isinstance(pce_factors, (str, bytes)):
        items: list[tuple[Any, Any]] = []
        for item in pce_factors:
            if isinstance(item, Mapping):
                items.extend(_factor_items(item))
            elif isinstance(item, (list, tuple)) and len(item) >= 2:
                items.append((item[0], item[1]))
        return items
    return []


def validate_pce_factors(pce_factors: Any = None) -> PceValidationResult:
    """Validate custom PCE factors and merge them over the defaults.

    The historical defaults are trusted as-is, including the current bicycle
    default of 0.0. User-provided overrides must be numeric and greater than
    zero; invalid values fall back to the default for that vehicle class.
    """

    defaults = get_default_pce_factors()
    factors = dict(defaults)
    issues: list[PceFactorIssue] = []
    seen: set[str] = set()

    for raw_vehicle_class, raw_factor in _factor_items(pce_factors):
        vehicle_class = str(raw_vehicle_class or "").strip()
        if not vehicle_class:
            continue
        if vehicle_class not in defaults:
            issues.append(
                PceFactorIssue(
                    "unknown_pce_vehicle_class",
                    "warning",
                    vehicle_class,
                    f"Unknown vehicle class ignored in PCE factors: {vehicle_class}",
                )
            )
            continue

        seen.add(vehicle_class)
        parsed = _coerce_factor(raw_factor)
        default_value = defaults[vehicle_class]
        if parsed is None:
            issues.append(
                PceFactorIssue(
                    "invalid_pce_factor",
                    "warning",
                    vehicle_class,
                    f"Invalid PCE factor for {vehicle_class}; using default {default_value:g}.",
                )
            )
            factors[vehicle_class] = default_value
            continue
        if parsed <= 0 and not _same_factor(parsed, default_value):
            issues.append(
                PceFactorIssue(
                    "invalid_pce_factor",
                    "warning",
                    vehicle_class,
                    f"PCE factor for {vehicle_class} must be greater than 0; using default {default_value:g}.",
                )
            )
            factors[vehicle_class] = default_value
            continue
        factors[vehicle_class] = parsed

    missing = [vehicle_class for vehicle_class in VEHICLE_CLASSES if vehicle_class not in seen and vehicle_class not in factors]
    for vehicle_class in missing:
        factors[vehicle_class] = defaults[vehicle_class]

    return PceValidationResult(factors=factors, issues=tuple(issues))


def normalize_pce_factors(pce_factors: Any = None) -> dict[str, float]:
    """Return validated PCE factors with missing known classes defaulted."""

    return validate_pce_factors(pce_factors).factors


def pce_factor_traceability_frame(pce_factors: Any = None) -> pd.DataFrame:
    """Return the PCE factors shown in exported workbooks."""

    defaults = get_default_pce_factors()
    factors = normalize_pce_factors(pce_factors)
    rows = []
    for vehicle_class in VEHICLE_CLASSES:
        factor = factors.get(vehicle_class, defaults[vehicle_class])
        rows.append(
            {
                "vehicle_class": vehicle_class,
                "pce_factor": factor,
                "source": "default" if _same_factor(factor, defaults[vehicle_class]) else "user_override",
            }
        )
    return pd.DataFrame(rows, columns=["vehicle_class", "pce_factor", "source"])


def pce_factors_equal(left: Any, right: Any) -> bool:
    """Return True when two PCE inputs resolve to the same selected factors."""

    left_factors = normalize_pce_factors(left)
    right_factors = normalize_pce_factors(right)
    return all(_same_factor(left_factors.get(vehicle_class), right_factors.get(vehicle_class)) for vehicle_class in VEHICLE_CLASSES)


def pce_factor_issue_frame(pce_factors: Any = None) -> pd.DataFrame:
    result = validate_pce_factors(pce_factors)
    return pd.DataFrame(
        [
            {
                "check": issue.check,
                "severity": issue.severity,
                "vehicle_class": issue.vehicle_class,
                "message": issue.message,
            }
            for issue in result.issues
        ],
        columns=["check", "severity", "vehicle_class", "message"],
    )


def apply_pce_factors(df: pd.DataFrame, pce_factors: dict[str, float] | None = None) -> pd.DataFrame:
    factors = normalize_pce_factors(PCE_FACTORS if pce_factors is None else pce_factors)
    result = df.copy()
    result["count"] = pd.to_numeric(result["count"], errors="coerce").fillna(0)
    result["pce_factor"] = result["vehicle_class"].map(factors)
    result["pcu"] = (result["count"] * result["pce_factor"]).fillna(0)
    return result


def add_pcu(df: pd.DataFrame, pce_factors: dict[str, float] | None = None) -> pd.DataFrame:
    return apply_pce_factors(df, pce_factors=pce_factors)
