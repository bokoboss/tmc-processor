"""Passenger car unit calculations."""

from __future__ import annotations

import pandas as pd

from .constants import PCE_FACTORS


def add_pcu(df: pd.DataFrame, pce_factors: dict[str, float] | None = None) -> pd.DataFrame:
    factors = pce_factors or PCE_FACTORS
    result = df.copy()
    result["count"] = pd.to_numeric(result["count"], errors="coerce").fillna(0)
    result["pce_factor"] = result["vehicle_class"].map(factors)
    result["pcu"] = result["count"] * result["pce_factor"]
    return result

