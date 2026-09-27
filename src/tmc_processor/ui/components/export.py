"""Export readiness/provenance display components."""

from __future__ import annotations

from collections.abc import Mapping

import streamlit as st


def render_export_backend(details: Mapping[str, object]) -> None:
    st.info(
        f"Requested: {details.get('requested', '')} · "
        f"Used: {details.get('used', '')} · "
        f"Fallback: {details.get('fallback', '') or 'none'}"
    )
