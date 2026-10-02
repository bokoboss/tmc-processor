"""UX-8 P1 stage hierarchy and disclosure checks."""

from __future__ import annotations

from pathlib import Path

import pytest
from streamlit.testing.v1 import AppTest

import app
from tmc_processor.excel_com_export import ExcelComStatus
from tmc_processor.ui.components.export import SAFE_PNG_TITLE, STANDARD_REPORT_TITLE, operator_report_label


ROOT = Path(__file__).resolve().parents[1]
DEMO = ROOT / "samples" / "demo"


@pytest.fixture(autouse=True)
def _without_excel_com(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(
        app,
        "_probe_excel_com_for_ui",
        lambda force=False: ExcelComStatus(available=False, reason="test", detail="", version=""),
    )


def _app() -> AppTest:
    at = AppTest.from_file(str(ROOT / "app.py"), default_timeout=60).run(timeout=60)
    assert not at.exception
    return at


def _stage(at: AppTest, label: str) -> AppTest:
    next(button for button in at.button if button.label == label and str(button.key).startswith("workflow_tab_")).click().run(timeout=60)
    assert not at.exception
    return at


def _batch(at: AppTest) -> AppTest:
    mode = next(widget for widget in at.radio if widget.key == "work_mode")
    mode.set_value(mode.options[1]).run(timeout=60)
    assert not at.exception
    return at


def _upload_demo(at: AppTest) -> AppTest:
    path = DEMO / "DEMO_TMC1_FourLeg.xlsx"
    at.file_uploader[0].set_value(
        (path.name, path.read_bytes(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    ).run(timeout=60)
    assert not at.exception
    return at


def _mapping_with_source(at: AppTest) -> AppTest:
    _upload_demo(at)
    return _stage(at, "Mapping")


def _expander(at: AppTest, label: str):
    return next(panel for panel in at.get("expander") if panel.label == label)


def test_single_data_groups_project_and_collapses_report_details() -> None:
    at = _app()
    keys = {widget.key for widget in at.text_input}
    assert {"project_name_input", "tmc_id_input", "survey_date_text_input", "survey_period_input"} <= keys
    details = _expander(at, "รายละเอียดสำหรับรายงาน")
    assert not details.proto.expanded
    detail_keys = {widget.key for widget in details.get("text_input")}
    assert {"weather_input", "responsible_party_input", "north_label_input", "north_road_input", "caption_text_input"} <= detail_keys
    assert any(widget.key == "show_u_turn_checkbox" for widget in details.get("checkbox"))


def test_survey_period_stays_on_data_and_persists_to_analyze() -> None:
    at = _app()
    period = next(widget for widget in at.text_input if widget.key == "survey_period_input")
    period.set_value("06.00 - 19.00").run(timeout=60)
    _stage(at, "Analyze")
    assert all(widget.key != "survey_period_input" for widget in at.text_input)
    assert at.session_state[app.SETUP_STATE_KEY]["survey_period"] == "06.00 - 19.00"


def test_single_mapping_leads_with_basic_editor_and_collapses_reuse() -> None:
    at = _mapping_with_source(_app())
    view = next(widget for widget in at.radio if widget.key == "mapping_editor_view_mode")
    assert view.value == "Basic"
    assert len(at.get("metric")) <= 4
    reuse = _expander(at, "นำเข้า / ใช้ Mapping ซ้ำ")
    advanced = _expander(at, "Advanced mapping controls")
    assert not reuse.proto.expanded and not advanced.proto.expanded
    assert {widget.key for widget in reuse.get("file_uploader")} == {"mapping_upload", "mapping_preset_upload"}


def test_single_mapping_preset_remains_usable_after_disclosure() -> None:
    at = _mapping_with_source(_app())
    preset = DEMO / "DEMO_TMC1_FourLeg.mapping.json"
    next(widget for widget in at.file_uploader if widget.key == "mapping_preset_upload").set_value(
        (preset.name, preset.read_bytes(), "application/json")
    ).run(timeout=60)
    assert not at.exception
    assert at.session_state["tmc_mapping_preset_apply_info"]["matched"] > 0


def test_advanced_mapping_scheme_choice_remains_available_before_commit() -> None:
    at = _mapping_with_source(_app())
    scheme = next(widget for widget in at.selectbox if widget.key == "movement_code_scheme_selector")
    assert not scheme.disabled
    scheme.set_value(scheme.options[-1]).run(timeout=60)
    assert not at.exception
    assert at.session_state["tmc_mapping_code_scheme"] == app.MOVEMENT_SCHEME_V2
    assert next(widget for widget in at.selectbox if widget.key == "movement_code_scheme_selector").value == app.MOVEMENT_SCHEME_V2


def test_single_analyze_shows_windows_and_collapses_advanced_settings() -> None:
    at = _stage(_app(), "Analyze")
    keys = {widget.key for widget in at.time_input}
    assert keys == {
        "am_peak_window_start_input", "am_peak_window_end_input",
        "pm_peak_window_start_input", "pm_peak_window_end_input",
    }
    assert all(widget.key != "survey_period_input" for widget in at.text_input)
    peak = next(panel for panel in at.get("expander") if panel.label.startswith("วิธีคำนวณ Peak"))
    pce = _expander(at, "ค่าเทียบเท่ารถยนต์นั่ง (PCE)")
    assert not peak.proto.expanded and not pce.proto.expanded
    assert next(widget for widget in peak.get("selectbox") if widget.key == "peak_mode_select").value == app.DEFAULT_PEAK_MODE


def test_single_review_keeps_technical_tables_in_one_closed_section() -> None:
    at = _mapping_with_source(_app())
    preset = DEMO / "DEMO_TMC1_FourLeg.mapping.json"
    next(widget for widget in at.file_uploader if widget.key == "mapping_preset_upload").set_value(
        (preset.name, preset.read_bytes(), "application/json")
    ).run(timeout=60)
    _stage(at, "Analyze")
    next(button for button in at.button if button.label == "วิเคราะห์ TMC").click().run(timeout=120)
    assert not at.exception
    assert at.session_state["active_workflow_tab"] == "Review"
    detail = _expander(at, "รายละเอียดทางเทคนิค")
    assert not detail.proto.expanded
    headings = "\n".join(str(item.value) for item in detail.get("markdown"))
    for label in ("Normalized Data", "Hourly Movement PCU", "Movement Summary", "Peak / PHF Data", "Movement Aggregation Audit", "Parser"):
        assert label in headings
    assert all(panel.label != "Normalized Data" for panel in at.get("expander"))
    assert any(button.label == "ยืนยันช่วง Peak" for button in at.button)


def test_batch_data_keeps_later_stage_controls_out() -> None:
    at = _batch(_app())
    assert any(widget.key == "survey_period_input" for widget in at.text_input)
    assert not _expander(at, "รายละเอียดสำหรับรายงาน Batch").proto.expanded
    assert all(widget.key != "peak_mode_select" for widget in at.selectbox)
    text = "\n".join(str(item.value) for item in at.markdown)
    assert "Shared PCE factors" not in text
    assert "Per-file Review" not in text
    assert "raw Excel ใน ZIP" not in text


def _batch_with_demo_uploads() -> AppTest:
    at = _batch(_app())
    mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    names = ("DEMO_TMC1_FourLeg.xlsx", "DEMO_TMC1_FourLeg_Day2.xlsx")
    next(widget for widget in at.file_uploader if str(widget.key).startswith("batch_raw_tmc_uploads")).set_value(
        [(name, (DEMO / name).read_bytes(), mime) for name in names]
    ).run(timeout=60)
    assert not at.exception
    assert len(at.session_state[app.BATCH_SOURCE_UPLOAD_STATE_KEY]) == 2
    return at


def test_batch_analyze_stays_disabled_without_mapping_preset() -> None:
    at = _stage(_batch_with_demo_uploads(), "Analyze")
    readiness = at.session_state[app.WORKFLOW_STATE_KEY][app.WORKFLOW_BATCH_MODE].readiness
    assert readiness.source and not readiness.mapping
    assert next(button for button in at.button if button.key == "analyze_batch_processing").disabled


def test_batch_analyze_receives_persisted_source_and_mapping_readiness() -> None:
    at = _stage(_batch_with_demo_uploads(), "Mapping")
    preset = DEMO / "DEMO_TMC1_FourLeg.mapping.json"
    next(widget for widget in at.file_uploader if str(widget.key).startswith("batch_mapping_preset_upload")).set_value(
        (preset.name, preset.read_bytes(), "application/json")
    ).run(timeout=60)
    assert not at.exception
    assert app.BATCH_SOURCE_UPLOAD_STATE_KEY in at.session_state
    assert app.BATCH_MAPPING_PRESET_UPLOAD_STATE_KEY in at.session_state
    at = _stage(at, "Analyze")
    readiness = at.session_state[app.WORKFLOW_STATE_KEY][app.WORKFLOW_BATCH_MODE].readiness
    assert readiness.source and readiness.mapping
    assert not next(button for button in at.button if button.key == "analyze_batch_processing").disabled


def test_batch_review_receives_saved_analysis_for_peak_and_qc(monkeypatch: pytest.MonkeyPatch) -> None:
    batch_contexts = []
    render_stage = app.render_workflow_stage

    def capture_review(stage: str, *, context):
        if stage in {"Review", "Export"} and not context.is_single_file_mode:
            batch_contexts.append(context)
        return render_stage(stage, context=context)

    monkeypatch.setattr(app, "render_workflow_stage", capture_review)
    at = _stage(_batch_with_demo_uploads(), "Mapping")
    preset = DEMO / "DEMO_TMC1_FourLeg.mapping.json"
    next(widget for widget in at.file_uploader if str(widget.key).startswith("batch_mapping_preset_upload")).set_value(
        (preset.name, preset.read_bytes(), "application/json")
    ).run(timeout=60)
    _stage(at, "Analyze")
    next(button for button in at.button if button.key == "analyze_batch_processing").click().run(timeout=120)
    assert not at.exception
    saved_analysis = at.session_state["tmc_batch_analysis_result"]
    assert len(saved_analysis.successful_items) == 2
    assert batch_contexts[-1].batch_analysis is saved_analysis
    assert batch_contexts[-1].batch_export_mode == at.session_state["tmc_batch_export_mode"]
    assert not batch_contexts[-1].batch_stale
    text = "\n".join(str(item.value) for item in at.markdown)
    assert "กรุณาวิเคราะห์ Batch ก่อนตรวจสอบ Peak" not in text
    assert any(widget.key == "tmc_batch_review_filter" for widget in at.radio)
    assert any(widget.label == "เลือกไฟล์สำหรับตรวจกราฟ" for widget in at.selectbox)
    assert any(str(button.key).startswith("batch_confirm_peak_review_") for button in at.button)
    assert "QC ข้อมูล" in text
    _stage(at, "Export")
    assert batch_contexts[-1].batch_analysis is saved_analysis
    assert batch_contexts[-1].batch_export_mode == at.session_state["tmc_batch_export_mode"]
    assert not batch_contexts[-1].batch_stale
    assert next(button for button in at.button if button.key == "generate_batch_zip").disabled


def test_batch_review_without_analysis_keeps_empty_state() -> None:
    at = _stage(_batch(_app()), "Review")
    assert at.session_state["tmc_batch_analysis_result"] is None
    text = "\n".join(str(item.value) for item in at.markdown)
    assert "กรุณาวิเคราะห์ Batch ก่อนตรวจสอบ Peak" in text
    assert not any(widget.key == "tmc_batch_review_filter" for widget in at.radio)


def test_batch_mapping_focuses_on_preset_and_sheet_matching() -> None:
    at = _stage(_batch(_app()), "Mapping")
    assert any(widget.key == "batch_mapping_preset_upload" for widget in at.file_uploader)
    assert any(panel.label == "สถานะ Sheet matching รายไฟล์" for panel in at.get("expander"))
    text = "\n".join(str(item.value) for item in at.markdown)
    assert "ค่า PCE พร้อมใช้งาน" not in text
    assert "Metadata รายไฟล์พร้อมใช้งาน" not in text


def test_batch_analyze_owns_pce_and_peak_settings_without_survey_period() -> None:
    at = _stage(_batch(_app()), "Analyze")
    assert len(at.time_input) == 4
    assert all(widget.key != "survey_period_input" for widget in at.text_input)
    assert not _expander(at, "ค่าเทียบเท่ารถยนต์นั่ง (PCE)").proto.expanded
    assert not next(panel for panel in at.get("expander") if panel.label.startswith("วิธีคำนวณ Peak")).proto.expanded
    assert any(button.label == "วิเคราะห์ Batch" for button in at.button)


@pytest.mark.parametrize("batch", [False, True])
def test_export_retains_ux7_product_choices(batch: bool) -> None:
    at = _app()
    if batch:
        _batch(at)
    _stage(at, "Export")
    copy = "\n".join(
        [str(item.value) for item in at.markdown]
        + [str(option) for widget in at.radio for option in widget.options]
    )
    assert STANDARD_REPORT_TITLE in copy
    assert "Advanced / Diagnostics" in copy
    assert operator_report_label(app.BATCH_SAFE_PNG_EXPORT_MODE if batch else app.SAFE_PNG_EXPORT_MODE) == SAFE_PNG_TITLE
    assert not at.exception
