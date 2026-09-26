"""Mapping workspace display components."""

from __future__ import annotations

from collections.abc import Mapping

import streamlit as st


def render_mapping_summary(summary: Mapping[str, object]) -> None:
    """Display mapping counts prepared by the application/domain layer."""

    st.metric("Mapping rows", int(summary.get("rows", 0)))
    st.caption(
        f"Included: {int(summary.get('included', 0))} · "
        f"Excluded: {int(summary.get('excluded', 0))} · "
        f"Peak included: {int(summary.get('peak_included', 0))}"
    )
