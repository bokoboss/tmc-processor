from __future__ import annotations

from datetime import time
from types import SimpleNamespace

import streamlit as st

import app
from tmc_processor.session import apply_session_to_state, build_project_session


def _seed_single_state(source_bytes: bytes = b"source-a") -> None:
    st.session_state.clear()
    app.initialize_setup_state_once("demo")
    st.session_state["mapping_table"] = [{"raw_sheet": "North", "movement_code": "NB"}]
    st.session_state["tmc_processed"] = {"result": SimpleNamespace()}
    st.session_state["tmc_output"] = {"workbook_bytes": b"old"}
    st.session_state["tmc_confirmed_am_peak_start"] = "08:00"
    st.session_state["tmc_confirmed_am_peak_end"] = "09:00"
    st.session_state["tmc_confirmed_pm_peak_start"] = "17:00"
    st.session_state["tmc_confirmed_pm_peak_end"] = "18:00"
    app._sync_single_workflow_from_state(
        source_bytes=source_bytes,
        source_file_name="demo.xlsx",
        export_mode=app.SAFE_PNG_EXPORT_MODE,
    )


def _sync(source_bytes: bytes = b"source-a"):
    return app._sync_single_workflow_from_state(
        source_bytes=source_bytes,
        source_file_name="demo.xlsx",
        export_mode=app.SAFE_PNG_EXPORT_MODE,
    )


def test_same_size_source_content_change_stales_single_analysis_and_export() -> None:
    _seed_single_state(b"abcd")

    transition = _sync(b"abce")

    assert transition.analysis_invalidated is True
    assert transition.review_invalidated is True
    assert transition.export_invalidated is True
    assert st.session_state["tmc_pce_results_stale"] is True
    assert "tmc_output" not in st.session_state


def test_mapping_semantic_change_stales_single_analysis_and_export() -> None:
    _seed_single_state()
    st.session_state["mapping_table"] = [{"raw_sheet": "North", "movement_code": "NS"}]

    transition = _sync()

    assert transition.analysis_invalidated is True
    assert st.session_state["tmc_pce_results_stale"] is True
    assert "tmc_output" not in st.session_state


def test_mapping_editor_and_view_only_changes_do_not_stale_single_state() -> None:
    _seed_single_state()
    st.session_state["mapping_editor_version"] = 99
    st.session_state["mapping_editor_view_mode"] = "Advanced"
    st.session_state["active_workflow_tab"] = "Review"
    st.session_state["expanded_panel"] = "diagnostics"

    transition = _sync()

    assert transition.changed_fields == ()
    assert "tmc_processed" in st.session_state
    assert "tmc_output" in st.session_state


def test_peak_search_change_stales_analysis_but_peak_decision_only_stales_export() -> None:
    _seed_single_state()
    st.session_state["peak_mode_select"] = "rolling_60min"
    st.session_state["am_peak_window_start_input"] = time(8, 0)
    peak_transition = _sync()

    assert peak_transition.analysis_invalidated is True
    assert st.session_state["tmc_pce_results_stale"] is True

    _seed_single_state()
    st.session_state["tmc_confirmed_am_peak_start"] = "09:00"
    st.session_state["tmc_confirmed_am_peak_end"] = "10:00"
    review_transition = _sync()

    assert review_transition.analysis_invalidated is False
    assert review_transition.review_updated is True
    assert review_transition.export_invalidated is True
    assert "tmc_processed" in st.session_state
    assert "tmc_output" not in st.session_state


def test_export_metadata_change_preserves_analysis_and_invalidates_artifact() -> None:
    _seed_single_state()
    st.session_state["tmc_title_input"] = "changed report title"

    transition = _sync()

    assert transition.analysis_invalidated is False
    assert transition.review_invalidated is False
    assert transition.export_invalidated is True
    assert "tmc_processed" in st.session_state
    assert "tmc_output" not in st.session_state


def test_project_session_with_unchanged_semantic_inputs_does_not_clear_artifacts() -> None:
    _seed_single_state()
    setup = app.get_current_setup_from_state("demo.xlsx")
    session = build_project_session(metadata=setup, directions=setup)
    previous = app._workflow_state_for_mode(app.WORKFLOW_SINGLE_MODE)

    apply_session_to_state(session, st.session_state)
    app.update_setup_state_from_widgets()
    app._hydrate_setup_widgets_from_state()
    transition = app._sync_single_workflow_from_state(
        source_revision=previous.revisions.source if previous is not None else None,
        source_file_name="demo.xlsx",
        export_mode=app.SAFE_PNG_EXPORT_MODE,
    )

    assert transition.changed_fields == ()
    assert "tmc_processed" in st.session_state
    assert "tmc_output" in st.session_state


class _Upload:
    def __init__(self, name: str, content: bytes) -> None:
        self.name = name
        self._content = content

    def getvalue(self) -> bytes:
        return self._content


def _seed_batch_state(source_bytes: bytes = b"source-a") -> tuple[_Upload, dict[str, object]]:
    st.session_state.clear()
    app.initialize_setup_state_once()
    upload = _Upload("demo.xlsx", source_bytes)
    preset = {
        "preset_name": "demo",
        "movement_code_scheme": "from_to",
        "mapping_rows": [{"raw_sheet": "North", "output_movement_code": "NB"}],
    }
    st.session_state["tmc_batch_file_metadata_table"] = [
        {"file_name": "demo.xlsx", "survey_date_text": "", "output_stem": "demo", "notes": ""}
    ]
    st.session_state["tmc_batch_analysis_result"] = SimpleNamespace(
        successful_items=[SimpleNamespace(confirmed_AM_peak="08:00-09:00", confirmed_PM_peak="17:00-18:00")]
    )
    st.session_state["tmc_batch_export_result"] = SimpleNamespace(package_bytes=b"zip")
    st.session_state["tmc_batch_confirmed_peaks"] = {
        "file_01_demo": {"AM": "08:00-09:00", "PM": "17:00-18:00"}
    }
    app._sync_batch_workflow_from_state(
        batch_uploads=[upload],
        mapping_preset=preset,
        movement_code_scheme="from_to",
        metadata_rows=st.session_state["tmc_batch_file_metadata_table"],
        export_mode=app.BATCH_SAFE_PNG_EXPORT_LABEL,
    )
    return upload, preset


def test_batch_source_content_change_stales_analysis_and_export() -> None:
    upload, preset = _seed_batch_state(b"abcd")
    changed_upload = _Upload(upload.name, b"abce")

    transition = app._sync_batch_workflow_from_state(
        batch_uploads=[changed_upload],
        mapping_preset=preset,
        movement_code_scheme="from_to",
        metadata_rows=st.session_state["tmc_batch_file_metadata_table"],
        export_mode=app.BATCH_SAFE_PNG_EXPORT_LABEL,
    )

    assert transition.analysis_invalidated is True
    assert st.session_state["tmc_batch_stale"] is True
    assert "tmc_batch_export_result" not in st.session_state


def test_batch_peak_decision_change_preserves_analysis_and_stales_export() -> None:
    upload, preset = _seed_batch_state()
    st.session_state["tmc_batch_confirmed_peaks"] = {
        "file_01_demo": {"AM": "09:00-10:00", "PM": "17:00-18:00"}
    }

    transition = app._sync_batch_workflow_from_state(
        batch_uploads=[upload],
        mapping_preset=preset,
        movement_code_scheme="from_to",
        metadata_rows=st.session_state["tmc_batch_file_metadata_table"],
        export_mode=app.BATCH_SAFE_PNG_EXPORT_LABEL,
    )

    assert transition.analysis_invalidated is False
    assert transition.review_updated is True
    assert st.session_state["tmc_batch_analysis_result"] is not None
    assert "tmc_batch_export_result" not in st.session_state


def test_batch_export_metadata_change_preserves_analysis_and_stales_export() -> None:
    upload, preset = _seed_batch_state()
    metadata_rows = [
        {"file_name": "demo.xlsx", "survey_date_text": "", "output_stem": "renamed", "notes": ""}
    ]
    st.session_state["tmc_batch_file_metadata_table"] = metadata_rows

    transition = app._sync_batch_workflow_from_state(
        batch_uploads=[upload],
        mapping_preset=preset,
        movement_code_scheme="from_to",
        metadata_rows=metadata_rows,
        export_mode=app.BATCH_SAFE_PNG_EXPORT_LABEL,
    )

    assert transition.analysis_invalidated is False
    assert transition.export_invalidated is True
    assert st.session_state["tmc_batch_analysis_result"] is not None
    assert "tmc_batch_export_result" not in st.session_state
