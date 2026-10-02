"""UX-8 P2 presentation derives from the existing workflow contract."""

from __future__ import annotations

from dataclasses import replace
import inspect
from pathlib import Path

import pytest
from streamlit.testing.v1 import AppTest

import app
from tmc_processor.excel_com_export import ExcelComStatus
from tmc_processor.session import CURRENT_SCHEMA_VERSION, session_from_json
from tmc_processor.ui.workflows.single import render_single_export
from tmc_processor.workflow_state import (
    WorkflowReadiness,
    WorkflowRevisions,
    WorkflowState,
    WorkflowTransition,
    readiness_after_transition,
    transition_workflow,
)


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


def _stage(at: AppTest, name: str) -> AppTest:
    next(button for button in at.button if button.key == f"workflow_tab_{app.CANONICAL_WORKFLOW_STAGES.index(name)}").click().run(timeout=60)
    assert not at.exception
    return at


def _status(at: AppTest) -> str:
    return next(item.value for item in at.get("caption") if str(item.value).startswith("สถานะขั้นตอน"))


def _single_source(at: AppTest) -> AppTest:
    path = DEMO / "DEMO_TMC1_FourLeg.xlsx"
    at.file_uploader[0].set_value(
        (path.name, path.read_bytes(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    ).run(timeout=60)
    assert not at.exception
    return at


def _batch(at: AppTest) -> AppTest:
    mode = next(widget for widget in at.radio if widget.key == "work_mode")
    mode.set_value(mode.options[1]).run(timeout=60)
    assert not at.exception
    return at


def test_project_session_controls_are_global_and_not_duplicated_in_export() -> None:
    assert "st.download_button" in inspect.getsource(app._render_project_session_section)
    assert "download_project_session_export" not in inspect.getsource(render_single_export)
    at = _app()
    for stage in app.CANONICAL_WORKFLOW_STAGES:
        _stage(at, stage)
        assert any(widget.key == "project_session_upload" and not widget.proto.disabled for widget in at.file_uploader)
        assert at.session_state["tmc_project_session_bytes"]


def test_batch_session_controls_are_visible_but_disabled_for_single_source_schema() -> None:
    at = _batch(_app())
    for stage in app.CANONICAL_WORKFLOW_STAGES:
        _stage(at, stage)
        assert any(widget.key == "project_session_upload_batch_disabled" and widget.proto.disabled for widget in at.file_uploader)
        assert any(button.key == "download_project_session_batch_disabled" and button.disabled for button in at.button)
        assert not any(widget.key == "project_session_upload" for widget in at.file_uploader)
        assert "Project Session ปัจจุบันรองรับงานไฟล์เดียว" in " ".join(item.value for item in at.get("caption"))


@pytest.mark.parametrize("mode", [app.WORKFLOW_SINGLE_MODE, app.WORKFLOW_BATCH_MODE])
def test_stage_statuses_follow_authoritative_readiness_and_invalidation(
    monkeypatch: pytest.MonkeyPatch, mode: str
) -> None:
    monkeypatch.setattr(app.st, "session_state", {})
    ready = WorkflowReadiness(source=True, mapping=True, analysis=True, review=True, export=True)
    state = WorkflowState(mode=mode, readiness=ready)
    assert all(value in {"พร้อม", "เสร็จแล้ว"} for value in app._workflow_stage_statuses(state).values())

    mapping_changed = readiness_after_transition(ready, WorkflowTransition(analysis_invalidated=True))
    after_mapping = app._workflow_stage_statuses(replace(state, readiness=mapping_changed))
    assert after_mapping["Analyze"] != "เสร็จแล้ว"
    assert after_mapping["Review"] != "เสร็จแล้ว"
    assert after_mapping["Export"] != "เสร็จแล้ว"

    peak_changed = readiness_after_transition(ready, WorkflowTransition(export_invalidated=True))
    after_peak = app._workflow_stage_statuses(replace(state, readiness=peak_changed))
    assert after_peak["Analyze"] == "เสร็จแล้ว"
    assert after_peak["Review"] == "เสร็จแล้ว"
    assert after_peak["Export"] == "พร้อมสร้าง"


def test_batch_first_export_is_ready_to_create_even_when_existing_stale_flag_is_set(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    state_values = {"tmc_batch_export_stale": True}
    monkeypatch.setattr(app.st, "session_state", state_values)
    state = WorkflowState(
        mode=app.WORKFLOW_BATCH_MODE,
        readiness=WorkflowReadiness(source=True, mapping=True, analysis=True, review=True),
    )
    assert app._workflow_stage_statuses(state)["Export"] == "พร้อมสร้าง"
    state_values["tmc_batch_export_signature"] = ("previous ZIP",)
    assert app._workflow_stage_statuses(state)["Export"] == "ต้องสร้างใหม่"


def test_navigation_cta_waits_for_readiness_and_only_changes_stage() -> None:
    at = _app()
    assert "Data: ต้องอัปโหลด" in _status(at)
    assert not any(button.key == "next_stage_data" for button in at.button)
    _single_source(at)
    assert "Data: พร้อม" in _status(at)
    next(button for button in at.button if button.key == "next_stage_data").click().run(timeout=60)
    assert at.session_state["active_workflow_tab"] == "Mapping"
    assert at.session_state.get("tmc_processed") is None
    assert at.session_state.get("tmc_output") is None
    assert not any(button.key == "next_stage_mapping" for button in at.button)


def test_single_mapping_cta_uses_ready_state_without_running_analysis() -> None:
    at = _stage(_single_source(_app()), "Mapping")
    preset = DEMO / "DEMO_TMC1_FourLeg.mapping.json"
    next(widget for widget in at.file_uploader if widget.key == "mapping_preset_upload").set_value(
        (preset.name, preset.read_bytes(), "application/json")
    ).run(timeout=60)
    assert "Mapping: พร้อม" in _status(at)
    next(button for button in at.button if button.key == "next_stage_mapping").click().run(timeout=60)
    assert at.session_state["active_workflow_tab"] == "Analyze"
    assert at.session_state.get("tmc_processed") is None
    assert not any(button.key == "next_stage_analyze" for button in at.button)


def test_batch_cta_uses_batch_readiness_without_running_analysis() -> None:
    at = _batch(_app())
    mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    names = ("DEMO_TMC1_FourLeg.xlsx", "DEMO_TMC1_FourLeg_Day2.xlsx")
    next(widget for widget in at.file_uploader if str(widget.key).startswith("batch_raw_tmc_uploads")).set_value(
        [(name, (DEMO / name).read_bytes(), mime) for name in names]
    ).run(timeout=60)
    assert "Data: พร้อม" in _status(at)
    next(button for button in at.button if button.key == "next_stage_data").click().run(timeout=60)
    assert at.session_state["active_workflow_tab"] == "Mapping"
    assert at.session_state.get("tmc_batch_analysis_result") is None
    preset = DEMO / "DEMO_TMC1_FourLeg.mapping.json"
    next(widget for widget in at.file_uploader if str(widget.key).startswith("batch_mapping_preset_upload")).set_value(
        (preset.name, preset.read_bytes(), "application/json")
    ).run(timeout=60)
    assert "Mapping: พร้อม" in _status(at)
    next(button for button in at.button if button.key == "next_stage_mapping").click().run(timeout=60)
    assert at.session_state["active_workflow_tab"] == "Analyze"
    assert at.session_state.get("tmc_batch_analysis_result") is None


def test_analyze_cta_only_navigates_to_review_after_current_result() -> None:
    at = _stage(_single_source(_app()), "Mapping")
    preset = DEMO / "DEMO_TMC1_FourLeg.mapping.json"
    next(widget for widget in at.file_uploader if widget.key == "mapping_preset_upload").set_value(
        (preset.name, preset.read_bytes(), "application/json")
    ).run(timeout=60)
    _stage(at, "Analyze")
    next(button for button in at.button if button.label == "วิเคราะห์ TMC").click().run(timeout=120)
    assert not at.exception
    assert "Analyze: เสร็จแล้ว" in _status(at)
    _stage(at, "Analyze")
    analysis = at.session_state.get("tmc_processed")
    next(button for button in at.button if button.key == "next_stage_analyze").click().run(timeout=60)
    assert at.session_state["active_workflow_tab"] == "Review"
    assert at.session_state.get("tmc_processed") is analysis
    assert at.session_state.get("tmc_output") is None
    assert not any(button.key == "next_stage_review" for button in at.button)


@pytest.mark.parametrize("mode", [app.WORKFLOW_SINGLE_MODE, app.WORKFLOW_BATCH_MODE])
@pytest.mark.parametrize(
    ("changed_field", "analysis_current", "review_current", "export_current"),
    [
        ("source", False, False, False),
        ("mapping", False, False, False),
        ("analysis_config", False, False, False),  # PCE or Peak-search change
        ("review_decision", True, True, False),  # confirmed Peak change
        ("export_config", True, True, False),  # report metadata or export setting
    ],
)
def test_display_follows_existing_revision_invalidation(
    monkeypatch: pytest.MonkeyPatch,
    mode: str,
    changed_field: str,
    analysis_current: bool,
    review_current: bool,
    export_current: bool,
) -> None:
    monkeypatch.setattr(app.st, "session_state", {})
    before = WorkflowRevisions(
        source="a", mapping="a", analysis_config="a", analysis_result="a", review_decision="a", export_config="a"
    )
    transition = transition_workflow(before, replace(before, **{changed_field: "b"}))
    readiness = readiness_after_transition(
        WorkflowReadiness(source=True, mapping=True, analysis=True, review=True, export=True), transition
    )
    status = app._workflow_stage_statuses(WorkflowState(mode=mode, readiness=readiness))
    assert (status["Analyze"] == "เสร็จแล้ว") is analysis_current
    assert (status["Review"] == "เสร็จแล้ว") is review_current
    assert (status["Export"] == "เสร็จแล้ว") is export_current


def test_project_session_round_trip_rebuilds_stage_presentation_without_ui_labels() -> None:
    at = _single_source(_app())
    _stage(at, "Mapping")
    preset = DEMO / "DEMO_TMC1_FourLeg.mapping.json"
    next(widget for widget in at.file_uploader if widget.key == "mapping_preset_upload").set_value(
        (preset.name, preset.read_bytes(), "application/json")
    ).run(timeout=60)
    at.run(timeout=60)
    saved = at.session_state["tmc_project_session_bytes"]
    loaded = session_from_json(saved).session
    assert loaded["schema_version"] == CURRENT_SCHEMA_VERSION
    assert "stage_statuses" not in loaded and "active_workflow_tab" not in loaded

    restored = _single_source(_app())
    next(widget for widget in restored.file_uploader if widget.key == "project_session_upload").set_value(
        ("saved.tmcproj.json", saved, "application/json")
    ).run(timeout=60)
    next(button for button in restored.button if button.key == "apply_project_session").click().run(timeout=60)
    assert not restored.exception
    assert restored.session_state.get("mapping_table")
    assert "Data: พร้อม" in _status(restored)
    assert "Mapping: พร้อม" in _status(restored)
    assert "Analyze: เสร็จแล้ว" not in _status(restored)
    assert restored.session_state.get("tmc_processed") is None
