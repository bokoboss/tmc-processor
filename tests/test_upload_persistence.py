from __future__ import annotations

from datetime import time
from types import SimpleNamespace

import pandas as pd
import streamlit as st

import app


class _Upload:
    def __init__(self, name: str, content: bytes) -> None:
        self.name = name
        self._content = content

    def getvalue(self) -> bytes:
        return self._content


def _reset() -> None:
    st.session_state.clear()


def _single_result() -> SimpleNamespace:
    return SimpleNamespace(
        peaks=pd.DataFrame(
            [
                {"period": "AM", "peak_start": time(8, 0), "peak_end": time(9, 0), "hourly_pcu": 100},
                {"period": "PM", "peak_start": time(17, 0), "peak_end": time(18, 0), "hourly_pcu": 120},
            ]
        )
    )


def _batch_preset() -> dict[str, object]:
    return {
        "preset_name": "Demo preset",
        "movement_code_scheme": "from_to",
        "mapping_rows": [{"raw_sheet": "North", "output_movement_code": "NB"}],
    }


def _seed_single_ready_state() -> _Upload:
    _reset()
    app.initialize_setup_state_once("source.xlsx")
    upload = _Upload("source.xlsx", b"source")
    st.session_state[app.SINGLE_SOURCE_UPLOAD_STATE_KEY] = app._upload_record(upload)
    st.session_state["mapping_table"] = [{"raw_sheet": "North", "movement_code": "NB"}]
    st.session_state["tmc_processed"] = {"result": _single_result()}
    st.session_state["tmc_output"] = {"workbook_bytes": b"old report"}
    st.session_state["tmc_confirmed_am_peak_start"] = "08:00"
    st.session_state["tmc_confirmed_am_peak_end"] = "09:00"
    st.session_state["tmc_confirmed_pm_peak_start"] = "17:00"
    st.session_state["tmc_confirmed_pm_peak_end"] = "18:00"
    app._sync_single_workflow_from_state(
        source_bytes=upload.getvalue(),
        source_file_name=upload.name,
        export_mode=app.SAFE_PNG_EXPORT_MODE,
    )
    return upload


def _seed_batch_ready_state() -> tuple[list[_Upload], dict[str, object]]:
    _reset()
    app.initialize_setup_state_once()
    uploads = [_Upload("one.xlsx", b"one"), _Upload("two.xlsx", b"two")]
    preset = _batch_preset()
    st.session_state[app.BATCH_SOURCE_UPLOAD_STATE_KEY] = [app._upload_record(upload) for upload in uploads]
    st.session_state[app.BATCH_MAPPING_PRESET_UPLOAD_STATE_KEY] = app._upload_record(_Upload("demo.json", b"preset"))
    st.session_state["tmc_batch_file_metadata_table"] = [
        {"file_name": upload.name, "survey_date_text": "", "output_stem": upload.name[:-5], "notes": ""}
        for upload in uploads
    ]
    item = SimpleNamespace(
        file_name="one.xlsx",
        folder_name="file_01_one",
        survey_date_text="",
        output_stem="one",
        notes="",
        workbook_bytes=b"one",
        confirmed_AM_peak="08:00-09:00",
        confirmed_PM_peak="17:00-18:00",
        status="success",
    )
    st.session_state["tmc_batch_analysis_result"] = SimpleNamespace(items=[item], successful_items=[item])
    st.session_state["tmc_batch_export_result"] = SimpleNamespace(package_bytes=b"old zip")
    st.session_state["tmc_batch_confirmed_peaks"] = {
        "file_01_one": {"AM": "08:00-09:00", "PM": "17:00-18:00"}
    }
    app._sync_batch_workflow_from_state(
        batch_uploads=uploads,
        mapping_preset=preset,
        movement_code_scheme="from_to",
        metadata_rows=st.session_state["tmc_batch_file_metadata_table"],
        export_mode=app.BATCH_SAFE_PNG_EXPORT_LABEL,
    )
    return uploads, preset


def test_single_upload_survives_canonical_stage_round_trip() -> None:
    _reset()
    upload = _Upload("source.xlsx", b"source")
    app._remember_upload(upload, app.SINGLE_SOURCE_UPLOAD_STATE_KEY)

    for stage in ("Data", "Mapping", "Analyze", "Data"):
        app.set_active_tab(stage)
        assert app._stored_upload(st.session_state.get(app.SINGLE_SOURCE_UPLOAD_STATE_KEY)).name == "source.xlsx"


def test_single_explicit_clear_removes_upload_and_invalidates_downstream() -> None:
    upload = _seed_single_ready_state()

    app._clear_single_source_upload()
    assert app.SINGLE_SOURCE_UPLOAD_STATE_KEY not in st.session_state

    transition = app._sync_single_workflow_from_state(
        source_bytes=None,
        source_file_name=None,
        export_mode=app.SAFE_PNG_EXPORT_MODE,
    )
    stored = app._workflow_state_for_mode(app.WORKFLOW_SINGLE_MODE)

    assert transition.analysis_invalidated is True
    assert stored is not None
    assert stored.readiness.source is False
    assert stored.readiness.analysis is False
    assert stored.readiness.review is False
    assert stored.readiness.export is False
    assert st.session_state["tmc_pce_results_stale"] is True
    assert "tmc_output" not in st.session_state
    assert upload.name == "source.xlsx"


def test_single_same_bytes_different_filename_updates_persisted_identity() -> None:
    _reset()

    app._remember_upload(_Upload("old.xlsx", b"same bytes"), app.SINGLE_SOURCE_UPLOAD_STATE_KEY)
    app._remember_upload(_Upload("new.xlsx", b"same bytes"), app.SINGLE_SOURCE_UPLOAD_STATE_KEY)

    stored = st.session_state[app.SINGLE_SOURCE_UPLOAD_STATE_KEY]
    assert stored["name"] == "new.xlsx"
    assert stored["data"] == b"same bytes"


def test_batch_uploads_survive_canonical_stage_round_trip() -> None:
    _reset()
    uploads = [_Upload("one.xlsx", b"one"), _Upload("two.xlsx", b"two")]
    app._remember_batch_uploads(uploads)

    for stage in ("Data", "Mapping", "Analyze", "Data"):
        app.set_active_tab(stage)
        stored = app._remember_batch_uploads(None)
        assert [upload.name for upload in stored] == ["one.xlsx", "two.xlsx"]


def test_batch_explicit_source_clear_removes_uploads_and_invalidates_downstream() -> None:
    uploads, preset = _seed_batch_ready_state()

    app._clear_batch_source_uploads()
    assert app.BATCH_SOURCE_UPLOAD_STATE_KEY not in st.session_state

    transition = app._sync_batch_workflow_from_state(
        batch_uploads=[],
        mapping_preset=preset,
        movement_code_scheme="from_to",
        metadata_rows=st.session_state["tmc_batch_file_metadata_table"],
        export_mode=app.BATCH_SAFE_PNG_EXPORT_LABEL,
    )
    stored = app._workflow_state_for_mode(app.WORKFLOW_BATCH_MODE)

    assert transition.analysis_invalidated is True
    assert stored is not None
    assert stored.readiness.source is False
    assert stored.readiness.analysis is False
    assert stored.readiness.review is False
    assert stored.readiness.export is False
    assert st.session_state["tmc_batch_stale"] is True
    assert "tmc_batch_export_result" not in st.session_state
    assert [upload.name for upload in uploads] == ["one.xlsx", "two.xlsx"]


def test_batch_same_bytes_different_filename_updates_persisted_identity() -> None:
    _reset()

    app._remember_batch_uploads([_Upload("old.xlsx", b"same bytes")])
    app._remember_batch_uploads([_Upload("new.xlsx", b"same bytes")])

    stored = st.session_state[app.BATCH_SOURCE_UPLOAD_STATE_KEY]
    assert stored[0]["name"] == "new.xlsx"
    assert stored[0]["data"] == b"same bytes"


def test_batch_mapping_preset_survives_navigation_and_explicit_clear_invalidates_downstream() -> None:
    uploads, preset = _seed_batch_ready_state()
    preset_upload = _Upload("demo.json", b"preset")

    for stage in ("Mapping", "Analyze", "Data", "Mapping"):
        app.set_active_tab(stage)
        assert app._stored_upload(st.session_state.get(app.BATCH_MAPPING_PRESET_UPLOAD_STATE_KEY)).name == "demo.json"

    app._clear_batch_mapping_preset_upload()
    assert app.BATCH_MAPPING_PRESET_UPLOAD_STATE_KEY not in st.session_state

    transition = app._sync_batch_workflow_from_state(
        batch_uploads=uploads,
        mapping_preset=None,
        movement_code_scheme="from_to",
        metadata_rows=st.session_state["tmc_batch_file_metadata_table"],
        export_mode=app.BATCH_SAFE_PNG_EXPORT_LABEL,
    )
    stored = app._workflow_state_for_mode(app.WORKFLOW_BATCH_MODE)

    assert transition.analysis_invalidated is True
    assert stored is not None
    assert stored.readiness.source is True
    assert stored.readiness.mapping is False
    assert stored.readiness.analysis is False
    assert stored.readiness.review is False
    assert stored.readiness.export is False
    assert st.session_state["tmc_batch_stale"] is True
    assert "tmc_batch_export_result" not in st.session_state
    assert preset_upload.name == "demo.json"
