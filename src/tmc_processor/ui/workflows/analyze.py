"""Analyze-stage rendering boundary."""

from __future__ import annotations

from collections.abc import Mapping

import streamlit as st


def render_analyze_stage(*, context: Mapping[str, object]) -> None:
    st.caption(str(context.get("description", "Run the established analysis pipeline.")))
