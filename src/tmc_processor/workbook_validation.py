"""Validation helpers for generated Excel workbook packages."""

from __future__ import annotations

import zipfile
from io import BytesIO
from pathlib import Path
from xml.etree import ElementTree as ET


def validate_workbook_xml(workbook: bytes | str | Path) -> tuple[str, ...]:
    """Parse every XML part in an .xlsx package and return validated part names.

    Raises ValueError when the workbook is missing the main worksheet XML or any
    XML part is not well-formed.
    """

    if isinstance(workbook, bytes):
        source = BytesIO(workbook)
    else:
        source = Path(workbook)

    with zipfile.ZipFile(source) as archive:
        names = archive.namelist()
        if "xl/worksheets/sheet1.xml" not in names:
            raise ValueError("Workbook package is missing xl/worksheets/sheet1.xml")

        xml_parts = tuple(name for name in names if name.endswith(".xml"))
        for name in xml_parts:
            try:
                ET.fromstring(archive.read(name))
            except ET.ParseError as exc:
                raise ValueError(f"Workbook XML part is not well-formed: {name}: {exc}") from exc
    return xml_parts
