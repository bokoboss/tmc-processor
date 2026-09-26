"""Review-stage rendering boundary."""

from __future__ import annotations

from collections.abc import Mapping

import streamlit as st

from tmc_processor.ui.components.peak import render_peak_card
from tmc_processor.ui.components.qc import render_qc_summary


def render_review_stage(*, context: Mapping[str, object]) -> None:
    counts = context.get("qc_counts")
    if isinstance(counts, Mapping):
        render_qc_summary(counts)
    peak = context.get("peak")
    if isinstance(peak, Mapping):
        render_peak_card(
            str(peak.get("title", "Peak")),
            str(peak.get("period", "")),
            str(peak.get("pcu", "")),
            str(peak.get("source", "")),
        )
    else:
        st.caption("Review QC evidence and confirm Peak decisions explicitly.")
