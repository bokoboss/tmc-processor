"""Secondary support surface, kept outside the canonical engineering stages."""

from __future__ import annotations

from pathlib import Path

import streamlit as st


SUPPORT_TITLE = "เลี้ยงชาเย็นคนทำโปรแกรม"
SUPPORT_QR_PATH = Path(__file__).resolve().parents[4] / "assets" / "support_qr.png"


def resolve_support_qr_path(asset_path: Path | None = None) -> Path | None:
    """Return the supplied QR asset when available, otherwise request a clean fallback."""

    candidate = SUPPORT_QR_PATH if asset_path is None else Path(asset_path)
    return candidate if candidate.is_file() else None


def render_support_view() -> None:
    st.title(SUPPORT_TITLE)
    st.markdown("TMC Processor เปิดให้ใช้งานฟรี")

    with st.container(border=True):
        text_column, qr_column = st.columns([1.5, 1], gap="large")
        with text_column:
            st.markdown(
                "ถ้าโปรแกรมนี้ช่วยลดเวลาจัดการข้อมูล TMC และงาน Excel ของคุณได้ "
                "และอยากช่วยสนับสนุนการพัฒนาต่อ จะเลี้ยงชาเย็นคนทำโปรแกรมสักแก้วก็ได้"
            )
            st.markdown("สนับสนุนหรือไม่ก็ใช้งานได้ครบทุกฟังก์ชันเหมือนเดิม")
        with qr_column:
            qr_path = resolve_support_qr_path()
            if qr_path is None:
                st.caption("ไม่มีไฟล์ QR สำหรับแสดงในสภาพแวดล้อมนี้")
            else:
                st.image(str(qr_path), width=280)
                st.markdown("**สแกน QR ได้ตามสะดวกเลย**")

    st.caption(
        "ในระยะนี้ TMC Processor ไม่มีการบันทึกไฟล์สมุดงานสำรวจที่อัปโหลดลงฐานข้อมูล "
        "หรือที่เก็บโครงการแบบถาวรโดยเจตนา ระหว่างใช้งาน ข้อมูลอาจอยู่ในหน่วยความจำของ "
        "session และการส่งออกอาจสร้างไฟล์ชั่วคราว ซึ่งระบบจะลบเมื่อจบการส่งออก"
    )
