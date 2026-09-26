"""Data-stage rendering boundary."""

from __future__ import annotations

from collections.abc import Mapping

import streamlit as st


def render_data_stage(*, context: Mapping[str, object]) -> None:
    st.caption(str(context.get("description", "Upload source workbooks and enter project setup.")))
