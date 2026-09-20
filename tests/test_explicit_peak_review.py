from __future__ import annotations

from datetime import time
from types import SimpleNamespace

import pandas as pd
import streamlit as st

import app
from tmc_processor.batch import BatchAnalysisItem, BatchAnalysisResult
from tmc_processor.session import apply_session_to_state


def _single_result() -> SimpleNamespace:
    return SimpleNamespace(
        peaks=pd.DataFrame(
            [
                {"period": "AM", "peak_start": time(7, 0), "peak_end": time(8, 0), "hourly_pcu": 100},
                {"period": "PM", "peak_start": time(17, 0), "peak_end": time(18, 0), "hourly_pcu": 120},
            ]
        )
    )


def _seed_single_state() -> None:
    st.session_state.clear()
    app.initialize_setup_state_once("demo")
    st.session_state["mapping_table"] = [{"raw_sheet": "North", "movement_code": "NB"}]
    st.session_state["tmc_processed"] = {"result": _single_result()}
    st.session_state["tmc_selected_pce_factors"] = app._current_pce_factors_from_state()


def _sync_single() -> object:
    return app._sync_single_workflow_from_state(
        source_bytes=b"source-a",
        source_file_name="demo.xlsx",
        export_mode=app.SAFE_PNG_EXPORT_MODE,
    )


def _seed_batch_state() -> None:
    st.session_state.clear()
    app.initialize_setup_state_once()
    item = BatchAnalysisItem(
        file_name="demo.xlsx",
        survey_date_text="",
        output_stem="demo",
        folder_name="file_01_demo",
        status="success",
        suggested_AM_peak="07:00-08:00",
        suggested_PM_peak="17:00-18:00",
    )
    st.session_state["tmc_batch_file_metadata_table"] = [
        {"file_name": "demo.xlsx", "survey_date_text": "", "output_stem": "demo", "notes": ""}
    ]
    st.session_state["tmc_batch_analysis_result"] = BatchAnalysisResult(items=[item])
    st.session_state["tmc_batch_confirmed_peaks"] = {}
    st.session_state["tmc_batch_draft_peaks"] = {}


def test_single_analysis_has_suggestions_but_not_review_confirmation() -> None:
    _seed_single_state()

    state = app._single_effective_peak_state()
    transition = _sync_single()

    assert state["suggested"]["am_peak_start"] == "07:00"
    assert app._confirmed_peaks_from_state() == {}
    assert state["ready"] is False
    assert transition.analysis_invalidated is False
    stored = app._workflow_state_for_mode(app.WORKFLOW_SINGLE_MODE)
    assert stored is not None
    assert stored.readiness.review is False
    assert stored.readiness.export is False


def test_single_explicit_confirmation_records_provenance_and_preserves_analysis() -> None:
    _seed_single_state()
    _sync_single()

    changed = app._confirm_single_peak_review(("07:00", "08:00"), ("17:00", "18:00"))
    transition = _sync_single()

    assert changed is True
    assert app._confirmed_peaks_from_state() == {
        "am_peak_start": "07:00",
        "am_peak_end": "08:00",
        "pm_peak_start": "17:00",
        "pm_peak_end": "18:00",
        "peak_selection_source": app.PEAK_SELECTION_USER_CONFIRMED,
    }
    assert transition.analysis_invalidated is False
    assert transition.review_updated is True
    stored = app._workflow_state_for_mode(app.WORKFLOW_SINGLE_MODE)
    assert stored is not None
    assert stored.readiness.analysis is True
    assert stored.readiness.review is True


def test_single_draft_change_does_not_replace_confirmed_or_effective_peak() -> None:
    _seed_single_state()
    _sync_single()
    app._confirm_single_peak_review(("07:00", "08:00"), ("17:00", "18:00"))
    _sync_single()
    st.session_state["tmc_output"] = {"workbook_bytes": b"old"}

    st.session_state["am_peak_period_select"] = "08:00-09:00"
    transition = _sync_single()
    state = app._single_effective_peak_state()

    assert transition.changed_fields == ()
    assert app._confirmed_peaks_from_state()["am_peak_start"] == "07:00"
    assert state["values"]["am_peak_start"] == "07:00"
    assert state["draft"]["am_peak_start"] == "08:00"
    assert "tmc_output" in st.session_state


def test_single_changed_confirmation_stales_export_without_reanalysis() -> None:
    _seed_single_state()
    _sync_single()
    app._confirm_single_peak_review(("07:00", "08:00"), ("17:00", "18:00"))
    _sync_single()
    st.session_state["tmc_output"] = {"workbook_bytes": b"old"}
    st.session_state["am_peak_period_select"] = "08:00-09:00"

    app._confirm_single_peak_review(("08:00", "09:00"), ("17:00", "18:00"))
    transition = _sync_single()

    assert transition.analysis_invalidated is False
    assert transition.review_updated is True
    assert transition.export_invalidated is True
    assert app._confirmed_peaks_from_state()["am_peak_start"] == "08:00"
    assert "tmc_processed" in st.session_state
    assert "tmc_output" not in st.session_state


def test_single_reconfirming_same_decision_has_no_semantic_churn() -> None:
    _seed_single_state()
    _sync_single()
    app._confirm_single_peak_review(("07:00", "08:00"), ("17:00", "18:00"))
    _sync_single()
    st.session_state["tmc_output"] = {"workbook_bytes": b"old"}

    changed = app._confirm_single_peak_review(("07:00", "08:00"), ("17:00", "18:00"))
    transition = _sync_single()

    assert changed is False
    assert transition.changed_fields == ()
    assert transition.analysis_invalidated is False
    assert transition.export_invalidated is False
    assert "tmc_output" in st.session_state


def test_confirmed_project_session_peak_remains_confirmed_after_reload() -> None:
    _seed_single_state()
    session = {
        "peaks": {
            "am_peak_start": "07:00",
            "am_peak_end": "08:00",
            "pm_peak_start": "17:00",
            "pm_peak_end": "18:00",
            "peak_selection_source": app.PEAK_SELECTION_USER_CONFIRMED,
        }
    }
    apply_session_to_state(session, st.session_state)

    _sync_single()

    state = app._single_effective_peak_state()
    assert state["ready"] is True
    assert state["source"] == app.PEAK_SELECTION_USER_CONFIRMED
    assert app._confirmed_peaks_from_state()["am_peak_start"] == "07:00"


def test_single_mapping_change_invalidates_explicit_confirmation() -> None:
    _seed_single_state()
    _sync_single()
    app._confirm_single_peak_review(("07:00", "08:00"), ("17:00", "18:00"))
    _sync_single()
    st.session_state["mapping_table"] = [{"raw_sheet": "North", "movement_code": "NS"}]

    transition = _sync_single()

    assert transition.analysis_invalidated is True
    assert transition.review_invalidated is True
    assert app._confirmed_peaks_from_state() == {}


def test_batch_analysis_starts_with_suggestions_and_no_confirmed_values() -> None:
    _seed_batch_state()

    analysis = st.session_state["tmc_batch_analysis_result"]
    assert analysis.successful_items[0].suggested_AM_peak == "07:00-08:00"
    assert analysis.successful_items[0].confirmed_AM_peak == ""
    assert analysis.successful_items[0].confirmed_PM_peak == ""
    assert app.reviewed_peak_values_complete(analysis) is False


def test_batch_draft_does_not_increment_confirmation_until_per_file_confirm() -> None:
    _seed_batch_state()
    item = st.session_state["tmc_batch_analysis_result"].successful_items[0]
    st.session_state["tmc_batch_draft_peaks"][item.folder_name] = {
        "AM": "08:00-09:00",
        "PM": "17:00-18:00",
    }

    assert item.confirmed_AM_peak == ""
    assert item.confirmed_PM_peak == ""
    assert not app.reviewed_peak_values_complete(st.session_state["tmc_batch_analysis_result"])

    changed = app._confirm_batch_peak_review(item.folder_name, "08:00-09:00", "17:00-18:00")

    assert changed is True
    assert item.confirmed_AM_peak == "08:00-09:00"
    assert item.confirmed_PM_peak == "17:00-18:00"
    assert app.reviewed_peak_values_complete(st.session_state["tmc_batch_analysis_result"])
