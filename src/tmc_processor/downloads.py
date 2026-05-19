"""Helpers for reliable Streamlit in-memory downloads."""

from __future__ import annotations

from io import BytesIO
import re


EXCEL_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
PNG_MIME = "image/png"


def download_buffer(data: bytes | bytearray | memoryview) -> BytesIO:
    """Return a seeked BytesIO buffer suitable for st.download_button."""

    buffer = BytesIO(bytes(data))
    buffer.seek(0)
    return buffer


def safe_workbook_filename(tmc_id: str | None, default: str = "tmc_processor_output.xlsx") -> str:
    """Build an ASCII workbook filename from a TMC id, falling back safely."""

    cleaned = re.sub(r"[^A-Za-z0-9_-]+", "_", str(tmc_id or "").strip()).strip("_")
    if not cleaned:
        return default
    return f"{cleaned[:60]}_output.xlsx"
