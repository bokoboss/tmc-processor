"""Export readiness/provenance display components."""

from __future__ import annotations

from collections.abc import Mapping

import streamlit as st


STANDARD_REPORT_TITLE = "รายงานมาตรฐาน — แนะนำ"
STANDARD_REPORT_DESCRIPTION = "รายงาน Excel ตามรูปแบบมาตรฐาน พร้อมกราฟและแผนผัง"
SAFE_PNG_TITLE = "รายงานสำรอง"
SAFE_PNG_DESCRIPTION = "ใช้รูปภาพสำหรับกราฟและแผนผังเพื่อความเข้ากันได้"
FALLBACK_MESSAGE = "ระบบใช้รูปแบบรายงานสำรองสำหรับการส่งออกครั้งนี้"


def operator_report_label(export_mode: str | None) -> str:
    if not export_mode:
        return "รอเลือกรูปแบบรายงาน"
    return STANDARD_REPORT_TITLE if "Excel Template Mode" in export_mode else SAFE_PNG_TITLE


def operator_fallback_message(technical_notice: str) -> str:
    return FALLBACK_MESSAGE if technical_notice else ""


def render_export_backend(details: Mapping[str, object]) -> None:
    st.info(
        f"Requested: {details.get('requested', '')} · "
        f"Used: {details.get('used', '')} · "
        f"Fallback: {details.get('fallback', '') or 'none'}"
    )
