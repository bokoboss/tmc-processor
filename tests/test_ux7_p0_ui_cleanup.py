"""Operator copy stays separate from export routing and technical diagnostics."""

from __future__ import annotations

from pathlib import Path

import pytest
from streamlit.testing.v1 import AppTest

import app
from tmc_processor.excel_com_export import ExcelComStatus
from tmc_processor.ui.components.export import (
    FALLBACK_MESSAGE,
    SAFE_PNG_TITLE,
    STANDARD_REPORT_TITLE,
    operator_fallback_message,
    operator_report_label,
)


ROOT = Path(__file__).resolve().parents[1]
DISPLAY_KINDS = ("markdown", "caption", "info", "warning", "success", "error", "button", "radio")


def _export_screen(monkeypatch: pytest.MonkeyPatch, *, available: bool, batch: bool = False) -> AppTest:
    monkeypatch.setattr(
        app,
        "_probe_excel_com_for_ui",
        lambda force=False: ExcelComStatus(
            available=available,
            reason="RAW_COM_REASON",
            detail="RAW_COM_DETAIL",
            version="99.9",
        ),
    )
    at = AppTest.from_file(str(ROOT / "app.py"), default_timeout=30).run(timeout=30)
    assert not at.exception
    if batch:
        mode = next(widget for widget in at.radio if widget.key == "work_mode")
        mode.set_value(mode.options[1]).run(timeout=30)
        # The existing Batch Standard mode adapter reruns while switching its
        # internal mode label. Isolate one render here; routing is tested below.
        at.session_state["active_workflow_tab"] = "Export"
        monkeypatch.setattr(app.st, "rerun", lambda: None)
        at.run(timeout=30)
    else:
        next(button for button in at.button if button.key == "workflow_tab_4").click().run(timeout=30)
    assert not at.exception
    return at


def _operator_text(at: AppTest) -> str:
    advanced = [panel for panel in at.get("expander") if panel.label.startswith("Advanced / Diagnostics")]
    assert advanced and all(not panel.proto.expanded for panel in advanced)
    hidden = {id(element) for panel in advanced for kind in DISPLAY_KINDS for element in panel.get(kind)}
    visible: list[str] = []
    for kind in DISPLAY_KINDS:
        for element in at.get(kind):
            if id(element) in hidden:
                continue
            visible.extend(str(getattr(element, field)) for field in ("label", "value") if hasattr(element, field) and kind != "radio")
            if kind == "radio":
                visible.append(str(element.label))
                visible.extend(element.options)
    return "\n".join(visible)


@pytest.mark.parametrize("available", [False, True])
def test_single_standard_export_hides_transport_and_keeps_diagnostics(monkeypatch: pytest.MonkeyPatch, available: bool) -> None:
    at = _export_screen(monkeypatch, available=available)
    visible = _operator_text(at)
    assert "Excel COM" not in visible
    assert "RAW_COM_REASON" not in visible
    assert "RAW_COM_DETAIL" not in visible
    assert "99.9" not in visible
    assert "Requested:" not in visible and "Used:" not in visible
    assert "ทดสอบ Excel COM" not in visible
    assert "Excel Template Mode" not in visible
    assert STANDARD_REPORT_TITLE in visible
    assert FALLBACK_MESSAGE not in visible

    diagnostics = "\n".join(
        str(element.value)
        for panel in at.get("expander") if panel.label.startswith("Advanced / Diagnostics")
        for element in panel.get("json")
    )
    assert "RAW_COM_REASON" in diagnostics
    assert "RAW_COM_DETAIL" in diagnostics
    assert "99.9" in diagnostics


def test_batch_standard_export_uses_product_copy_and_keeps_internal_mode_gate(monkeypatch: pytest.MonkeyPatch) -> None:
    at = _export_screen(monkeypatch, available=False, batch=True)
    visible = _operator_text(at)
    assert "Excel COM" not in visible
    assert "RAW_COM_REASON" not in visible
    assert "โหมดส่งออกพร้อม" not in visible
    assert FALLBACK_MESSAGE not in visible
    assert STANDARD_REPORT_TITLE in visible
    diagnostic_text = "\n".join(
        str(element.value)
        for panel in at.get("expander") if panel.label.startswith("Advanced / Diagnostics")
        for element in panel.get("json")
    )
    assert "RAW_COM_REASON" in diagnostic_text


def test_presentation_helpers_do_not_change_standard_or_safe_png_decision() -> None:
    ready = ExcelComStatus(available=True, reason="ready")
    unavailable = ExcelComStatus(available=False, reason="RAW_COM_REASON")
    native = app._standard_report_decision(ready)
    without_com = app._standard_report_decision(unavailable)
    fallback = app._standard_report_decision(unavailable, template_compatible=False)

    assert native.use_template_report_layout is True
    assert native.use_excel_com_native_charts is False
    assert native.use_ooxml_native_template is True
    assert native.fallback_notice == ""
    assert without_com.use_template_report_layout is True
    assert without_com.use_ooxml_native_template is True
    assert without_com.fallback_notice == ""
    assert fallback.use_template_report_layout is False
    assert fallback.use_excel_com_native_charts is False
    assert "incompatible" in fallback.fallback_notice
    assert operator_fallback_message(native.fallback_notice) == ""
    assert operator_fallback_message(fallback.fallback_notice) == FALLBACK_MESSAGE
    assert operator_report_label(app.EXCEL_TEMPLATE_EXPORT_MODE) == STANDARD_REPORT_TITLE
    assert operator_report_label(app.SAFE_PNG_EXPORT_MODE) == SAFE_PNG_TITLE
    assert app._single_export_mode_options(unavailable) == [app.SAFE_PNG_EXPORT_MODE]
    assert app._single_export_mode_options(ready) == [app.EXCEL_TEMPLATE_EXPORT_MODE, app.SAFE_PNG_EXPORT_MODE]
