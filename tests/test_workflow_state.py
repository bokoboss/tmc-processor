from __future__ import annotations

from datetime import time
from types import SimpleNamespace

import pandas as pd
import streamlit as st

import app
from tmc_processor.batch import BatchAnalysisItem, BatchAnalysisResult, BatchResult
from tmc_processor.excel_com_export import ExcelComStatus


def _excel_ready() -> ExcelComStatus:
    return ExcelComStatus(available=True, reason="ok")


def _processed_result() -> SimpleNamespace:
    peaks = pd.DataFrame(
        [
            {"period": "AM", "peak_start": time(8, 0), "peak_end": time(9, 0), "hourly_pcu": 100},
            {"period": "PM", "peak_start": time(17, 0), "peak_end": time(18, 0), "hourly_pcu": 120},
        ]
    )
    return SimpleNamespace(peaks=peaks)


def _reset_state() -> None:
    st.session_state.clear()


def _set_batch_analyzed_state() -> None:
    _reset_state()
    st.session_state["tmc_batch_file_metadata_table"] = [{"file_name": "a.xlsx", "survey_date_text": "", "output_stem": "a", "notes": ""}]
    st.session_state["tmc_batch_input_signature"] = ("same",)
    st.session_state["tmc_batch_analysis_result"] = BatchAnalysisResult(
        items=[
            BatchAnalysisItem(
                file_name="a.xlsx",
                survey_date_text="",
                output_stem="a",
                folder_name="file_01_a",
                status="success",
                workbook_bytes=b"x",
                confirmed_AM_peak="08:00-09:00",
                confirmed_PM_peak="17:00-18:00",
            )
        ]
    )
    st.session_state["tmc_batch_confirmed_peaks"] = {"file_01_a": {"AM": "08:00-09:00", "PM": "17:00-18:00"}}


def _single_ready_state() -> dict[str, object]:
    _reset_state()
    st.session_state["mapping_table"] = [{"raw_sheet": "North", "movement_code": "NB"}]
    st.session_state["tmc_processed"] = {"result": _processed_result()}
    return app.derive_single_workflow_state("demo.xlsx", app.SAFE_PNG_EXPORT_MODE, _excel_ready())


def test_single_processed_effective_peaks_export_ready_not_completed() -> None:
    state = _single_ready_state()

    assert state["steps"][3] == "completed"
    assert state["steps"][4] == "ready"
    assert state["steps"][5] == "ready"
    assert ("ส่งออก", "พร้อมสร้างรายงาน", "success") in state["summary"]


def test_active_tab_state_persists_after_export_mode_change() -> None:
    _reset_state()
    st.session_state["tmc_output"] = {"workbook_bytes": b"x"}
    app.set_active_tab("ส่งออก")

    changed = app.apply_single_export_mode_change(app.EXCEL_TEMPLATE_EXPORT_MODE, app.SAFE_PNG_EXPORT_MODE)

    assert changed is True
    assert app.get_active_tab() == "ส่งออก"
    assert st.session_state["report_export_mode"] == app.EXCEL_TEMPLATE_EXPORT_MODE
    assert "tmc_output" not in st.session_state


def test_process_action_can_set_active_tab_to_peak_review() -> None:
    _reset_state()

    app.set_active_tab("ตรวจ Peak")

    assert app.get_active_tab() == "ตรวจ Peak"


def test_export_action_keeps_active_tab_as_export() -> None:
    _reset_state()

    app.set_active_tab("ส่งออก")

    assert app.get_active_tab() == "ส่งออก"


def test_invalid_active_tab_falls_back_to_default_setup() -> None:
    _reset_state()
    st.session_state["active_workflow_tab"] = "ไม่ใช่แท็บ"

    assert app.get_active_tab() == "ตั้งค่า"
    assert st.session_state["active_workflow_tab"] == "ตั้งค่า"


def test_work_mode_switch_does_not_reset_valid_active_tab() -> None:
    _reset_state()
    app.set_active_tab("ส่งออก")

    st.session_state["work_mode"] = "ประมวลผลหลายไฟล์"

    assert app.get_active_tab() == "ส่งออก"


def test_workflow_tab_choices_match_required_labels() -> None:
    assert app.workflow_tab_choices() == ["ตั้งค่า", "กำหนดทิศทาง", "ตรวจ Peak", "ส่งออก", "ตรวจสอบข้อมูล"]


def test_single_generated_report_marks_export_completed() -> None:
    _reset_state()
    st.session_state["mapping_table"] = [{"raw_sheet": "North", "movement_code": "NB"}]
    st.session_state["tmc_processed"] = {"result": _processed_result()}
    st.session_state["tmc_output"] = {"workbook_bytes": b"x"}

    state = app.derive_single_workflow_state("demo.xlsx", app.SAFE_PNG_EXPORT_MODE, _excel_ready())

    assert state["steps"][5] == "completed"
    assert ("ส่งออก", "สร้างแล้ว", "success") in state["summary"]


def test_peak_change_clears_export_and_returns_workflow_to_export_ready() -> None:
    _reset_state()
    st.session_state["mapping_table"] = [{"raw_sheet": "North", "movement_code": "NB"}]
    st.session_state["tmc_processed"] = {"result": _processed_result()}
    st.session_state["tmc_output"] = {"workbook_bytes": b"old"}

    assert app._clear_single_export() is True

    state = app.derive_single_workflow_state("demo.xlsx", app.SAFE_PNG_EXPORT_MODE, _excel_ready())

    assert state["steps"][5] == "ready"
    assert ("ส่งออก", "พร้อมสร้างรายงาน", "success") in state["summary"]


def test_single_processed_result_peak_summary_uses_suggested_not_no_results() -> None:
    state = _single_ready_state()

    peak_chip = next(item for item in state["summary"] if item[0] == "Peak")
    assert peak_chip == ("Peak", "ใช้ค่าแนะนำ", "success")


def test_single_confirmed_peak_summary_uses_release_status_wording() -> None:
    _reset_state()
    st.session_state["mapping_table"] = [{"raw_sheet": "North", "movement_code": "NB"}]
    st.session_state["tmc_processed"] = {"result": _processed_result()}
    st.session_state["tmc_confirmed_am_peak_start"] = "09:00"
    st.session_state["tmc_confirmed_am_peak_end"] = "10:00"
    st.session_state["tmc_confirmed_pm_peak_start"] = "17:00"
    st.session_state["tmc_confirmed_pm_peak_end"] = "18:00"
    st.session_state["am_peak_period_select"] = "09:00-10:00"

    state = app.derive_single_workflow_state("demo.xlsx", app.SAFE_PNG_EXPORT_MODE, _excel_ready())

    peak_chip = next(item for item in state["summary"] if item[0] == "Peak")
    assert peak_chip == ("Peak", "กำหนดแล้ว", "success")


def test_batch_analyzed_peaks_export_ready_not_completed() -> None:
    _reset_state()
    st.session_state["tmc_batch_file_metadata_table"] = [{"file_name": "a.xlsx", "output_stem": "a"}]
    st.session_state["tmc_batch_input_signature"] = ("same",)
    st.session_state["tmc_batch_analysis_result"] = BatchAnalysisResult(
        items=[
            BatchAnalysisItem(
                file_name="a.xlsx",
                survey_date_text="",
                output_stem="a",
                folder_name="a",
                status="success",
                confirmed_AM_peak="08:00-09:00",
                confirmed_PM_peak="17:00-18:00",
            )
        ]
    )

    state = app.derive_batch_workflow_state(uploaded_count=1, batch_mapping_ready=True, batch_signature=("same",))

    assert state["steps"][3] == "completed"
    assert state["steps"][4] == "ready"
    assert state["steps"][5] == "ready"
    assert ("ส่งออก", "พร้อมสร้าง Batch ZIP", "success") in state["summary"]


def test_batch_export_mode_change_does_not_reset_active_tab() -> None:
    _set_batch_analyzed_state()
    st.session_state["tmc_batch_export_result"] = BatchResult(package_bytes=b"zip")
    app.set_active_tab("ส่งออก")

    changed = app.apply_batch_export_mode_change(app.BATCH_EXCEL_TEMPLATE_EXPORT_LABEL, app.BATCH_SAFE_PNG_EXPORT_LABEL)

    assert changed is True
    assert app.get_active_tab() == "ส่งออก"
    assert st.session_state["tmc_batch_export_mode"] == app.BATCH_EXCEL_TEMPLATE_EXPORT_LABEL
    assert not st.session_state.get("tmc_batch_stale", False)
    assert st.session_state["tmc_batch_export_stale"] is True
    assert "tmc_batch_export_result" not in st.session_state


def test_batch_analyzed_change_export_mode_keeps_analysis_valid_and_export_stale() -> None:
    _set_batch_analyzed_state()
    st.session_state["tmc_batch_export_result"] = BatchResult(package_bytes=b"zip")

    changed = app.apply_batch_export_mode_change(app.BATCH_EXCEL_TEMPLATE_EXPORT_LABEL, app.BATCH_SAFE_PNG_EXPORT_LABEL)
    state = app.derive_batch_workflow_state(uploaded_count=1, batch_mapping_ready=True, batch_signature=("same",))

    assert changed is True
    assert not st.session_state.get("tmc_batch_stale", False)
    assert st.session_state["tmc_batch_analysis_result"] is not None
    assert st.session_state["tmc_batch_export_stale"] is True
    assert ("Batch Analysis", "วิเคราะห์แล้ว", "success") in state["summary"]
    assert ("Peak", "กำหนดแล้ว 1/1 ไฟล์", "success") in state["summary"]
    assert ("ส่งออก", "ต้องสร้าง ZIP ใหม่", "warning") in state["summary"]


def test_batch_analyzed_change_output_stem_keeps_analysis_valid_and_export_stale() -> None:
    _set_batch_analyzed_state()
    st.session_state["tmc_batch_export_result"] = BatchResult(package_bytes=b"zip")
    st.session_state["tmc_batch_file_metadata_table"] = [{"file_name": "a.xlsx", "survey_date_text": "", "output_stem": "renamed", "notes": ""}]

    app._sync_batch_analysis_metadata_from_state()
    app._mark_batch_export_stale_now()

    item = st.session_state["tmc_batch_analysis_result"].items[0]
    assert not st.session_state.get("tmc_batch_stale", False)
    assert st.session_state["tmc_batch_export_stale"] is True
    assert item.output_stem == "renamed"
    assert item.folder_name == "file_01_renamed"
    assert "tmc_batch_export_result" not in st.session_state


def test_batch_analyzed_change_confirmed_peak_keeps_analysis_valid_and_export_stale() -> None:
    _set_batch_analyzed_state()
    st.session_state["tmc_batch_export_result"] = BatchResult(package_bytes=b"zip")

    item = st.session_state["tmc_batch_analysis_result"].items[0]
    item.confirmed_AM_peak = "09:00-10:00"
    st.session_state["tmc_batch_confirmed_peaks"] = {"file_01_a": {"AM": "09:00-10:00", "PM": "17:00-18:00"}}
    app._mark_batch_export_stale_now()

    assert not st.session_state.get("tmc_batch_stale", False)
    assert item.confirmed_AM_peak == "09:00-10:00"
    assert st.session_state["tmc_batch_export_stale"] is True
    assert "tmc_batch_export_result" not in st.session_state


def test_batch_analyzed_change_pce_marks_analysis_stale() -> None:
    _set_batch_analyzed_state()

    is_stale = app._mark_batch_stale_if_inputs_changed(("pce", "changed"))

    assert is_stale is True
    assert st.session_state["tmc_batch_stale"] is True
    assert st.session_state["tmc_batch_export_stale"] is True


def test_batch_analyzed_change_mapping_preset_marks_analysis_stale() -> None:
    _set_batch_analyzed_state()

    is_stale = app._mark_batch_stale_if_inputs_changed(("mapping-preset", "changed"))

    assert is_stale is True
    assert st.session_state["tmc_batch_stale"] is True


def test_single_file_export_mode_change_keeps_processed_and_peaks_valid() -> None:
    _reset_state()
    st.session_state["tmc_processed"] = {"result": _processed_result()}
    st.session_state["tmc_output"] = {"workbook_bytes": b"x"}
    st.session_state["tmc_confirmed_am_peak_start"] = "08:00"
    st.session_state["tmc_confirmed_am_peak_end"] = "09:00"
    st.session_state["tmc_confirmed_pm_peak_start"] = "17:00"
    st.session_state["tmc_confirmed_pm_peak_end"] = "18:00"

    changed = app.apply_single_export_mode_change(app.EXCEL_TEMPLATE_EXPORT_MODE, app.SAFE_PNG_EXPORT_MODE)

    assert changed is True
    assert st.session_state["tmc_processed"]["result"] is not None
    assert app._confirmed_peaks_from_state()["am_peak_start"] == "08:00"
    assert "tmc_output" not in st.session_state


def test_batch_analyze_action_can_set_active_tab_to_peak_review() -> None:
    _reset_state()

    app.set_active_tab("ตรวจ Peak")

    assert app.get_active_tab() == "ตรวจ Peak"


def test_batch_zip_marks_export_completed() -> None:
    _reset_state()
    st.session_state["tmc_batch_file_metadata_table"] = [{"file_name": "a.xlsx", "output_stem": "a"}]
    st.session_state["tmc_batch_input_signature"] = ("same",)
    st.session_state["tmc_batch_analysis_result"] = BatchAnalysisResult(
        items=[
            BatchAnalysisItem(
                file_name="a.xlsx",
                survey_date_text="",
                output_stem="a",
                folder_name="a",
                status="success",
                confirmed_AM_peak="08:00-09:00",
                confirmed_PM_peak="17:00-18:00",
            )
        ]
    )
    st.session_state["tmc_batch_export_result"] = BatchResult(package_bytes=b"zip")

    state = app.derive_batch_workflow_state(uploaded_count=1, batch_mapping_ready=True, batch_signature=("same",))

    assert state["steps"][5] == "completed"
    assert ("ส่งออก", "สร้าง Batch ZIP แล้ว", "success") in state["summary"]


def test_stale_batch_process_warning_and_export_not_ready() -> None:
    _reset_state()
    st.session_state["tmc_batch_file_metadata_table"] = [{"file_name": "a.xlsx", "output_stem": "a"}]
    st.session_state["tmc_batch_input_signature"] = ("old",)
    st.session_state["tmc_batch_analysis_result"] = BatchAnalysisResult(
        items=[
            BatchAnalysisItem(
                file_name="a.xlsx",
                survey_date_text="",
                output_stem="a",
                folder_name="a",
                status="success",
                confirmed_AM_peak="08:00-09:00",
                confirmed_PM_peak="17:00-18:00",
            )
        ]
    )

    state = app.derive_batch_workflow_state(uploaded_count=1, batch_mapping_ready=True, batch_signature=("new",))

    assert state["steps"][3] == "warning"
    assert state["steps"][5] == "pending"
    assert ("ส่งออก", "ยังไม่พร้อม", "neutral") in state["summary"]
