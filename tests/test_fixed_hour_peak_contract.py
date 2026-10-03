from datetime import time
import json

import pandas as pd
import pytest

from tmc_processor.constants import DEFAULT_PEAK_MODE, PEAK_MODE_OPTIONS
from tmc_processor.peaks import confirmed_peak_phf, detect_peak_phf
from tmc_processor.application.state import confirm_single_peaks, confirm_batch_peak
from tmc_processor.session import session_from_json, apply_session_to_state


def test_product_peak_options_are_fixed_only():
    assert DEFAULT_PEAK_MODE == "fixed_hourly"
    assert PEAK_MODE_OPTIONS == ["fixed_hourly"]


@pytest.mark.parametrize("start,end", [("08:15", "09:15"), ("08:30", "09:30"),
    ("08:00", "09:15"), ("08:00", "10:00"), ("25:00", "26:00"),
    ("08:00junk", "09:00"), ("08:00:01", "09:00:01"), ("23:00", "00:00")])
def test_confirmation_rejects_unsupported_period_even_without_data(start, end):
    with pytest.raises(ValueError, match="whole-hour"):
        confirmed_peak_phf(pd.DataFrame(), {"AM": (start, end)})


@pytest.mark.parametrize("confirm", [
    lambda state: confirm_single_peaks(state, "08:15-09:15", "17:00-18:00"),
    lambda state: confirm_batch_peak(state, "demo", "08:15-09:15", "17:00-18:00"),
])
def test_application_confirmation_rejects_before_state_mutation(confirm):
    state = {"sentinel": "untouched"}
    with pytest.raises(ValueError, match="whole-hour"):
        confirm(state)
    assert state == {"sentinel": "untouched"}


def test_rolling_analysis_is_rejected():
    with pytest.raises(ValueError):
        detect_peak_phf(pd.DataFrame(), peak_mode="rolling_60min")


def test_legacy_session_preserves_settings_but_requires_fresh_analysis():
    raw = {"schema_version": 1, "metadata": {"project_name": "Legacy survey"},
        "pce_factors": {"MC": 0.7}, "mapping": {"rows": [{"raw_sheet": "North", "movement_code": "NS"}]},
        "peaks": {"peak_mode": "rolling_60min", "am_peak_start": "08:15", "am_peak_end": "09:15",
        "pm_peak_start": "17:00", "pm_peak_end": "18:00", "peak_selection_source": "user_confirmed"}}
    loaded = session_from_json(json.dumps(raw))
    assert loaded.session["metadata"]["project_name"] == "Legacy survey"
    assert loaded.session["mapping"]["rows"][0]["raw_sheet"] == "North"
    assert loaded.session["pce_factors"]["MC"] == 0.7
    assert loaded.session["peaks"]["peak_mode"] == "fixed_hourly"
    assert not loaded.session["peaks"].get("am_peak_start")
    assert any("not reused" in warning and "Analyze" in warning for warning in loaded.warnings)
    state = {"tmc_processed": {"old": True}, "tmc_output": b"old",
        "tmc_application_confirmed_peaks": {"am_peak_start": "08:15"},
        "tmc_confirmed_am_peak_start": "08:15", "tmc_loaded_confirmed_peaks": {"old": True}}
    apply_session_to_state(loaded.session, state)
    assert state["peak_mode_select"] == "fixed_hourly"
    assert "tmc_processed" not in state and "tmc_output" not in state
    assert "tmc_application_confirmed_peaks" not in state
    assert not state.get("tmc_loaded_confirmed_peaks")


def test_auto_peak_uses_whole_hours_inside_unaligned_search_window():
    data = pd.DataFrame([{"time_start": time(7 + i // 4, (i % 4) * 15),
        "include_in_peak": True, "pcu": [1, 50, 50, 50, 50, 1, 1, 1][i]}
        for i in range(8)])
    peaks = detect_peak_phf(data, windows={"AM": ("07:15", "09:00")})
    assert peaks.iloc[0]["peak_start"] == time(8, 0)
    assert peaks.iloc[0]["peak_end"] == time(9, 0)


@pytest.mark.parametrize("period", [("07:00", "08:00"), ("11:00", "12:00"), ("15:00", "16:00")])
def test_supported_confirmation_and_same_reconfirmation(period):
    state = {}
    label = "-".join(period)
    assert confirm_single_peaks(state, label, "17:00-18:00") is True
    assert confirm_single_peaks(state, label, "17:00-18:00") is False


@pytest.mark.parametrize("mode", ["fixed_hourly", "rolling_60min"])
def test_incompatible_session_always_clears_both_decisions(mode):
    loaded = session_from_json(json.dumps({"peaks": {"peak_mode": mode,
        "am_peak_start": "08:15", "am_peak_end": "09:15",
        "pm_peak_start": "17:00", "pm_peak_end": "18:00"}}))
    assert loaded.session["peak_reanalysis_required"]
    assert loaded.session["peaks"]["pm_peak_start"] == ""
    assert loaded.warnings


def test_fixed_hour_session_has_no_migration_and_preserves_decisions():
    peaks = {"peak_mode": "fixed_hourly", "am_peak_start": "08:00", "am_peak_end": "09:00",
             "pm_peak_start": "17:00", "pm_peak_end": "18:00", "peak_selection_source": "user_confirmed"}
    loaded = session_from_json(json.dumps({"peaks": peaks}))
    assert not loaded.warnings
    assert loaded.session["peaks"] == peaks
    state = {}
    apply_session_to_state(loaded.session, state)
    assert state["tmc_loaded_confirmed_peaks"]["am_peak_start"] == "08:00"


def test_review_options_do_not_inject_legacy_peaks():
    import app
    hourly = pd.DataFrame({"Time": ["08:00-09:00", "08:15-09:15", "Total"]})
    peaks = pd.DataFrame([{"period": "AM", "peak_start": time(8, 15), "peak_end": time(9, 15)}])
    assert app._hourly_interval_options(hourly, peaks) == [("08:00-09:00", "08:00", "09:00")]


@pytest.mark.parametrize("batch", [False, True])
def test_analyze_ui_has_no_peak_mode_selector_even_with_stale_rolling_state(batch, monkeypatch):
    from streamlit.testing.v1 import AppTest
    from pathlib import Path
    import app
    from tmc_processor.excel_com_export import ExcelComStatus
    monkeypatch.setattr(app, "_probe_excel_com_for_ui", lambda force=False: ExcelComStatus(
        available=False, reason="contract test without optional Excel COM", detail="", version=""))
    at = AppTest.from_file(str(Path(__file__).resolve().parents[1] / "app.py"), default_timeout=60).run()
    if batch:
        at.radio(key="work_mode").set_value(at.radio(key="work_mode").options[1]).run()
    at.session_state["peak_mode_select"] = "rolling_60min"
    at.button(key="workflow_tab_2").click().run()
    assert not at.exception
    assert all(widget.key != "peak_mode_select" for widget in at.selectbox)
    assert at.session_state[app.SETUP_STATE_KEY]["peak_mode"] == "fixed_hourly"


def test_direct_analysis_apis_cannot_activate_rolling():
    from tmc_processor.pipeline import process_tmc, process_tmc_dry_run_v2
    from tmc_processor.batch import analyze_batch_files, generate_batch_zip_from_reviewed_peaks, BatchAnalysisResult
    for analyze in (process_tmc, process_tmc_dry_run_v2):
        with pytest.raises(ValueError, match="fixed_hourly"):
            analyze({}, pd.DataFrame(), {}, peak_mode="rolling_60min")
    with pytest.raises(ValueError, match="fixed_hourly"):
        analyze_batch_files([], peak_mode="rolling_60min")
    with pytest.raises(ValueError, match="fixed_hourly"):
        generate_batch_zip_from_reviewed_peaks(BatchAnalysisResult(items=[]), peak_mode="rolling_60min")


def test_direct_export_rejects_stale_non_hourly_result():
    from tmc_processor.exporter import export_workbook
    peaks = pd.DataFrame([{"period": "AM", "peak_start": time(8, 15), "peak_end": time(9, 15)}])
    with pytest.raises(ValueError, match="whole-hour"):
        export_workbook({}, *[pd.DataFrame() for _ in range(6)], peaks)


def test_session_builder_does_not_truncate_invalid_peak_to_an_hour():
    from tmc_processor.session import build_project_session, ProjectSessionError
    with pytest.raises(ProjectSessionError, match="whole-hour"):
        build_project_session(peak_settings={"am_peak_start": "08:00:01", "am_peak_end": "09:00:01"})


def test_legacy_migration_invalidates_authoritative_workflow_directly():
    from tmc_processor.workflow_state import WorkflowState, WorkflowReadiness, WorkflowRevisions
    previous = WorkflowState(mode="single", revisions=WorkflowRevisions(analysis_result="old", review_decision="old"),
        readiness=WorkflowReadiness(source=True, mapping=True, analysis=True, review=True, export=True))
    state = {"tmc_workflow_state": {"single": previous}}
    apply_session_to_state({"peaks": {"peak_mode": "rolling_60min"}}, state)
    current = state["tmc_workflow_state"]["single"]
    assert current.readiness == WorkflowReadiness(source=True, mapping=True)
    assert current.revisions.analysis_result is None and current.revisions.review_decision is None
