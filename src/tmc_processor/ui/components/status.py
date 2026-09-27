"""Small status/readiness display primitives."""

from __future__ import annotations

from html import escape
from collections.abc import Iterable

import streamlit as st


def render_status_card(label: str, status: str, note: str = "", *, tone: str = "neutral") -> None:
    safe_tone = tone if tone in {"neutral", "green", "amber", "red", "blue"} else "neutral"
    note_html = f'<div class="tmc-status-note">{escape(note)}</div>' if note else ""
    st.markdown(
        f'<div class="tmc-status-card tmc-status-{safe_tone}">'
        f'<div class="tmc-status-label">{escape(label)}</div>'
        f'<div class="tmc-status-value">{escape(status)}</div>{note_html}</div>',
        unsafe_allow_html=True,
    )


def render_readiness_checklist(items: Iterable[tuple[str, bool, str]]) -> None:
    rows = []
    for label, ready, detail in items:
        icon = "✓" if ready else "!"
        status = "Ready" if ready else "Needs review"
        detail_html = f'<div class="tmc-check-detail">{escape(detail)}</div>' if detail else ""
        rows.append(
            f'<div class="tmc-check-item {"tmc-check-ready" if ready else "tmc-check-warn"}">'
            f'<span class="tmc-check-icon">{icon}</span><span class="tmc-check-body">'
            f'<div class="tmc-check-label">{escape(label)}</div>'
            f'<div class="tmc-check-status">{status}</div>{detail_html}</span></div>'
        )
    st.markdown('<div class="tmc-checklist">' + "".join(rows) + "</div>", unsafe_allow_html=True)
