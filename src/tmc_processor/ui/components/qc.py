"""QC evidence display components.

The component accepts already-computed QC rows and only renders them.
"""

from __future__ import annotations

from collections.abc import Mapping

import streamlit as st


def render_qc_summary(counts: Mapping[str, int], *, label: str = "QC status") -> None:
    errors = int(counts.get("error", 0))
    warnings = int(counts.get("warning", 0))
    info = int(counts.get("info", 0))
    if errors:
        status, tone = "Errors found", "red"
    elif warnings or info:
        status, tone = f"{warnings} warnings / {info} info", "amber"
    else:
        status, tone = "No QC issues", "green"
    st.markdown(
        f'<div class="tmc-status-card tmc-status-{tone}">'
        f'<div class="tmc-status-label">{label}</div>'
        f'<div class="tmc-status-value">{status}</div></div>',
        unsafe_allow_html=True,
    )
