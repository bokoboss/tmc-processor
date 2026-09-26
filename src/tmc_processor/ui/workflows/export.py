"""Export-stage rendering boundary."""

from __future__ import annotations

from collections.abc import Mapping

import streamlit as st

from tmc_processor.ui.components.export import render_export_backend


def render_export_stage(*, context: Mapping[str, object]) -> None:
    details = context.get("backend")
    if isinstance(details, Mapping):
        render_export_backend(details)
    else:
        st.caption("Choose the export backend and download the reviewed artifact.")
