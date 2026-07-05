from __future__ import annotations

from pathlib import Path
import shutil

import pytest

import tmc_processor.excel_com_export as excel_com_export


class _FakeCell:
    HasFormula = False
    Formula = ""
    Value = ""


class _FakeWorksheet:
    def __init__(self) -> None:
        self.cells = {"A3": _FakeCell()}

    def Range(self, cell_ref: str):
        return self.cells[cell_ref]


def test_excel_com_summary_label_uses_effective_peak_metadata() -> None:
    worksheet = _FakeWorksheet()

    excel_com_export._write_effective_peak_label(
        worksheet,
        {
            "am_peak_start": "17:00",
            "am_peak_end": "18:00",
            "pm_peak_start": "17:00",
            "pm_peak_end": "18:00",
            "peak_selection_source": "user_confirmed",
        },
    )

    assert worksheet.cells["A3"].Value == "Peak used for export: AM 17:00-18:00; PM 17:00-18:00 (user_confirmed)"


def test_export_with_excel_com_copy_permission_error_tells_user_to_close_excel(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    source = tmp_path / "template.xlsx"
    target = tmp_path / "report.xlsx"
    source.write_bytes(b"template")

    monkeypatch.setattr(excel_com_export, "require_excel_com", lambda: None)
    monkeypatch.setattr(excel_com_export, "_import_com_modules_with_repair", lambda: (object(), object()))

    def locked_copy(_source: Path, _target: Path) -> None:
        raise PermissionError("locked")

    monkeypatch.setattr(shutil, "copy2", locked_copy)

    with pytest.raises(PermissionError) as exc_info:
        excel_com_export.export_with_excel_com(
            source,
            target,
            template_map={},
            report_data={},
            metadata={},
            chart_source_data={},
        )

    message = str(exc_info.value)
    assert "ปิดไฟล์ Excel" in message
    assert str(target.resolve()) in message
    assert isinstance(exc_info.value.__cause__, PermissionError)
