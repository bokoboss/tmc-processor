"""Peak evidence display; Peak selection/calculation remains in domain code."""

from __future__ import annotations

from html import escape

import streamlit as st


def render_peak_card(title: str, period_label: str, pcu: str, source: str) -> None:
    confirmed = source == "user_confirmed"
    kind = "tmc-peak-confirmed" if confirmed else "tmc-peak-suggested"
    badge = "Confirmed" if confirmed else "Suggested"
    pcu_text = f"{pcu} PCU" if pcu else "No PCU value"
    st.markdown(
        f'<div class="tmc-card tmc-peak-card {kind}">'
        f'<div class="tmc-peak-topline"><div class="tmc-card-label">{escape(title)}</div>'
        f'<div class="tmc-peak-badge">{badge}</div></div>'
        f'<div class="tmc-peak-time">{escape(period_label or "No period")}</div>'
        f'<div class="tmc-peak-pcu">{escape(pcu_text)}</div></div>',
        unsafe_allow_html=True,
    )
