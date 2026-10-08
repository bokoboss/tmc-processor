"""Compact support dialog, kept outside the canonical engineering stages."""

from __future__ import annotations

from pathlib import Path

import streamlit as st


SUPPORT_TITLE = "เลี้ยงชาเย็นคนทำโปรแกรม"
SUPPORT_BODY = """TMC Processor เปิดให้ใช้งานฟรีครับ

ถ้าโปรแกรมนี้ช่วยลดเวลาจัดการข้อมูล TMC และงาน Excel ของคุณได้
และอยากช่วยสนับสนุนการพัฒนาต่อ
สามารถเลี้ยงชาเย็นคนทำโปรแกรมได้ตามสะดวกครับ

การสนับสนุนเป็นความสมัครใจ และไม่มีผลต่อฟังก์ชันการใช้งานของโปรแกรม"""
PRIVACY_DISCLOSURE = (
    "ในระยะนี้ TMC Processor ไม่มีการบันทึกไฟล์สมุดงานสำรวจที่อัปโหลดลงฐานข้อมูล "
    "หรือที่เก็บโครงการแบบถาวรโดยเจตนา ระหว่างใช้งาน ข้อมูลอาจอยู่ในหน่วยความจำของ "
    "session และการส่งออกอาจสร้างไฟล์ชั่วคราว ซึ่งระบบจะลบเมื่อจบการส่งออก"
)
SUPPORT_QR_PATH = Path(__file__).resolve().parents[4] / "assets" / "support_qr.png"


def resolve_support_qr_path(asset_path: Path | None = None) -> Path | None:
    """Return the supplied QR asset when available, otherwise request a clean fallback."""

    candidate = SUPPORT_QR_PATH if asset_path is None else Path(asset_path)
    return candidate if candidate.is_file() else None


def render_support_view() -> None:
    """Render the compact support dialog contents without modifying session state."""

    st.markdown(SUPPORT_BODY)
    qr_path = resolve_support_qr_path()
    if qr_path is None:
        st.caption("ไม่มีไฟล์ QR สำหรับแสดงในสภาพแวดล้อมนี้")
    else:
        _left, center, _right = st.columns([1, 2, 1], gap="small")
        with center:
            st.image(str(qr_path), width=220)


@st.dialog(SUPPORT_TITLE, width="small")
def _support_dialog() -> None:
    render_support_view()


def open_support_dialog() -> None:
    """Open the support dialog without routing away from the active workflow."""

    _support_dialog()
