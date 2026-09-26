"""Mapping-stage rendering boundary."""

from __future__ import annotations

from collections.abc import Mapping

import streamlit as st

from tmc_processor.ui.components.mapping import render_mapping_summary


def render_mapping_stage(*, context: Mapping[str, object]) -> None:
    summary = context.get("summary")
    if isinstance(summary, Mapping):
        render_mapping_summary(summary)
    else:
        st.caption("Define and review the physical movement mapping.")
