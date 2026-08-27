from __future__ import annotations

from datetime import time
from types import SimpleNamespace

import pytest
import pandas as pd
import streamlit as st

import app
from tmc_processor.excel_com_export import ExcelComStatus
from tmc_processor.session import apply_session_to_state, build_project_session
from tmc_processor.workflow_state import WorkflowReadiness, WorkflowRevisions, WorkflowState


def _base_workflow_revisions() -> WorkflowRevisions:
    return WorkflowRevisions(
        source="source-a",
        mapping="mapping-a",
        analysis_config="analysis-a",
        analysis_result="result-a",
        review_decision="review-a",
        export_config="export-a",
    )


@pytest.mark.parametrize(
    ("changes", "expected_readiness"),
    [
        (
            {"source": "source-b"},
            WorkflowReadiness(source=True, mapping=True, analysis=False, review=False, export=False),
        ),
        (
            {"mapping": "mapping-b"},
            WorkflowReadiness(source=True, mapping=True, analysis=False, review=False, export=False),
        ),
        (
            {"analysis_config": "analysis-b"},
            WorkflowReadiness(source=True, mapping=True, analysis=False, review=False, export=False),
        ),
        (
            {"analysis_result": "result-b"},
            WorkflowReadiness(source=True, mapping=True, analysis=True, review=False, export=False),
        ),
        (
            {"review_decision": "review-b"},
            WorkflowReadiness(source=True, mapping=True, analysis=True, review=True, export=False),
        ),
        (
            {"export_config": "export-b"},
            WorkflowReadiness(source=True, mapping=True, analysis=True, review=True, export=False),
        ),
        (
            {},
            WorkflowReadiness(source=True, mapping=True, analysis=True, review=True, export=True),
        ),
    ],
)
def test_stored_workflow_readiness_matches_transition(
    changes: dict[str, str],
    expected_readiness: WorkflowReadiness,
) -> None:
    st.session_state.clear()
    base = _base_workflow_revisions()
    full_readiness = WorkflowReadiness(source=True, mapping=True, analysis=True, review=True, export=True)
    app._sync_workflow_contract(app.WORKFLOW_SINGLE_MODE, base, full_readiness)

    app._sync_workflow_contract(
        app.WORKFLOW_SINGLE_MODE,
        base.with_updates(**changes),
        full_readiness,
    )

    stored = app._workflow_state_for_mode(app.WORKFLOW_SINGLE_MODE)
    assert stored is not None
    assert stored.readiness == expected_readiness


def _seed_single_state(source_bytes: bytes = b"source-a", mapping_code: str = "NB") -> None:
    st.session_state.clear()
    app.initialize_setup_state_once("demo")
    st.session_state["mapping_table"] = [{"raw_sheet": "North", "movement_code": mapping_code}]
    st.session_state["tmc_processed"] = {"result": SimpleNamespace()}
    st.session_state["tmc_output"] = {"workbook_bytes": b"old"}
    st.session_state["tmc_confirmed_am_peak_start"] = "08:00"
    st.session_state["tmc_confirmed_am_peak_end"] = "09:00"
    st.session_state["tmc_confirmed_pm_peak_start"] = "17:00"
    st.session_state["tmc_confirmed_pm_peak_end"] = "18:00"
    st.session_state["tmc_selected_pce_factors"] = app._current_pce_factors_from_state()
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


def _single_shell_state() -> dict[str, object]:
    return app.derive_single_workflow_state(
        "demo.xlsx",
        app.SAFE_PNG_EXPORT_MODE,
        ExcelComStatus(available=True, reason="ok"),
    )


def _batch_shell_state() -> dict[str, object]:
    return app.derive_batch_workflow_state(
        uploaded_count=1,
        batch_mapping_ready=True,
        batch_signature=("legacy",),
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


def test_mapping_editor_change_updates_stored_readiness_in_same_adapter_flow() -> None:
    _seed_single_state(mapping_code="NE")
    base_mapping = pd.DataFrame(st.session_state["mapping_table"])
    edited_mapping = base_mapping.copy()
    edited_mapping.loc[0, "movement_code"] = "NS"
    st.session_state["mapping_table"] = app._merge_mapping_editor_result(
        base_mapping,
        edited_mapping,
    ).to_dict("records")

    transition = app._sync_single_workflow_from_state(
        source_bytes=b"source-a",
        source_file_name="demo.xlsx",
        export_mode=app.SAFE_PNG_EXPORT_MODE,
    )

    stored = app._workflow_state_for_mode(app.WORKFLOW_SINGLE_MODE)
    assert transition.analysis_invalidated is True
    assert st.session_state["tmc_pce_results_stale"] is True
    assert "tmc_output" not in st.session_state
    assert stored is not None
    assert stored.readiness.analysis is False
    assert stored.readiness.review is False
    assert stored.readiness.export is False
    shell = _single_shell_state()
    assert shell["readiness"] == {"source": True, "mapping": True, "analysis": False, "review": False, "export": False}
    assert shell["steps"][2:] == ["active", "pending", "pending"]


def test_pce_editor_change_updates_stored_readiness_in_same_adapter_flow() -> None:
    _seed_single_state()
    changed_pce = dict(st.session_state["tmc_selected_pce_factors"])
    changed_pce["MC"] = 0.5

    app._store_selected_pce_factors(changed_pce)
    app._sync_workflow_after_pce_editor(
        is_single_file_mode=True,
        source_bytes=b"source-a",
        source_file_name="demo.xlsx",
        export_mode=app.SAFE_PNG_EXPORT_MODE,
    )

    stored = app._workflow_state_for_mode(app.WORKFLOW_SINGLE_MODE)
    assert stored is not None
    assert st.session_state["tmc_pce_results_stale"] is True
    assert stored.readiness.analysis is False
    assert stored.readiness.review is False
    assert stored.readiness.export is False
    shell = _single_shell_state()
    assert shell["readiness"]["analysis"] is False
    assert shell["readiness"]["review"] is False
    assert shell["readiness"]["export"] is False
    assert shell["steps"][2:] == ["active", "pending", "pending"]


def test_single_shell_prefers_stored_readiness_over_conflicting_legacy_state() -> None:
    _seed_single_state()
    st.session_state.pop("tmc_pce_results_stale", None)
    app._store_workflow_state(
        WorkflowState(
            mode=app.WORKFLOW_SINGLE_MODE,
            readiness=WorkflowReadiness(source=True, mapping=True, analysis=False, review=False, export=False),
        )
    )

    shell = _single_shell_state()

    assert shell["readiness"] == {"source": True, "mapping": True, "analysis": False, "review": False, "export": False}
    assert shell["steps"][2:] == ["active", "pending", "pending"]


def test_top_status_bar_prefers_stored_source_and_mapping_readiness(monkeypatch: pytest.MonkeyPatch) -> None:
    _seed_single_state()
    app._store_workflow_state(
        WorkflowState(
            mode=app.WORKFLOW_SINGLE_MODE,
            readiness=WorkflowReadiness(source=False, mapping=False),
        )
    )
    rendered: list[str] = []
    monkeypatch.setattr(app.st, "markdown", lambda body, **_: rendered.append(body))

    app._render_top_status_bar(
        is_single_file_mode=True,
        uploaded_name="demo.xlsx",
        uploaded_count=0,
        batch_mapping_ready=False,
        export_mode=app.SAFE_PNG_EXPORT_MODE,
        excel_com_status=ExcelComStatus(available=True, reason="ok"),
    )

    assert rendered
    assert "ยังไม่มีไฟล์สำรวจ" in rendered[0]
    assert "ยังไม่พร้อม" in rendered[0]


def test_pce_override_survives_widget_rehydration_without_spurious_transition() -> None:
    _seed_single_state()
    default_pce = dict(st.session_state["tmc_selected_pce_factors"])
    changed_pce = dict(default_pce)
    changed_pce["MC"] = 0.5

    app._store_selected_pce_factors(changed_pce)
    app._sync_workflow_after_pce_editor(
        is_single_file_mode=True,
        source_bytes=b"source-a",
        source_file_name="demo.xlsx",
        export_mode=app.SAFE_PNG_EXPORT_MODE,
    )
    st.session_state["tmc_pce_results_stale"] = False
    st.session_state["tmc_processed"] = {"result": SimpleNamespace(), "pce_factors": changed_pce}
    st.session_state["tmc_output"] = {"workbook_bytes": b"new"}
    st.session_state["tmc_confirmed_am_peak_start"] = "08:00"
    st.session_state["tmc_confirmed_am_peak_end"] = "09:00"
    st.session_state["tmc_confirmed_pm_peak_start"] = "17:00"
    st.session_state["tmc_confirmed_pm_peak_end"] = "18:00"
    app._sync_workflow_after_pce_editor(
        is_single_file_mode=True,
        source_bytes=b"source-a",
        source_file_name="demo.xlsx",
        export_mode=app.SAFE_PNG_EXPORT_MODE,
    )
    st.session_state["tmc_pce_results_stale"] = False
    st.session_state["tmc_output"] = {"workbook_bytes": b"reprocessed"}
    # Reprocessing may clear legacy review keys while applying the result transition;
    # restore the confirmed Peak as the normal post-review state before reopening Settings.
    st.session_state["tmc_confirmed_am_peak_start"] = "08:00"
    st.session_state["tmc_confirmed_am_peak_end"] = "09:00"
    st.session_state["tmc_confirmed_pm_peak_start"] = "17:00"
    st.session_state["tmc_confirmed_pm_peak_end"] = "18:00"
    app._sync_workflow_after_pce_editor(
        is_single_file_mode=True,
        source_bytes=b"source-a",
        source_file_name="demo.xlsx",
        export_mode=app.SAFE_PNG_EXPORT_MODE,
    )

    # Streamlit can retain stale widget values from before the canonical table changed.
    st.session_state["pce_factors_editor_0"] = app._pce_factor_records(default_pce)
    st.session_state["pce_factors_editor_1"] = app._pce_factor_records(default_pce)
    rehydrated = app._current_pce_factors_from_state()
    assert rehydrated["MC"] == 0.5
    traceability = app.pce_factor_traceability_frame(rehydrated)
    mc_row = traceability.loc[traceability["vehicle_class"] == "MC"].iloc[0]
    assert mc_row["source"] == "user_override"

    app._store_selected_pce_factors(rehydrated)
    transition = app._sync_workflow_after_pce_editor(
        is_single_file_mode=True,
        source_bytes=b"source-a",
        source_file_name="demo.xlsx",
        export_mode=app.SAFE_PNG_EXPORT_MODE,
    )

    assert transition.changed_fields == ()
    assert st.session_state["tmc_pce_results_stale"] is False
    assert "tmc_output" in st.session_state
    stored = app._workflow_state_for_mode(app.WORKFLOW_SINGLE_MODE)
    assert stored is not None
    assert stored.readiness.analysis is True
    assert stored.readiness.review is True
    assert stored.readiness.export is True
    shell = _single_shell_state()
    assert shell["readiness"] == {"source": True, "mapping": True, "analysis": True, "review": True, "export": True}
    assert shell["steps"][2:] == ["completed", "ready", "completed"]


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
    shell = _single_shell_state()
    assert shell["readiness"] == {"source": True, "mapping": True, "analysis": True, "review": True, "export": True}
    assert shell["steps"][2:] == ["completed", "ready", "completed"]


def test_canonical_stage_navigation_does_not_invalidate_single_workflow() -> None:
    _seed_single_state()
    previous = app._workflow_state_for_mode(app.WORKFLOW_SINGLE_MODE)
    assert previous is not None

    for stage in app.workflow_stages_for_mode(app.WORKFLOW_SINGLE_MODE):
        app.set_active_tab(stage)

    transition = _sync()
    current = app._workflow_state_for_mode(app.WORKFLOW_SINGLE_MODE)

    assert transition.changed_fields == ()
    assert current is not None
    assert current.revisions == previous.revisions
    assert "tmc_processed" in st.session_state
    assert "tmc_output" in st.session_state


def test_rehydrate_analyze_setup_widgets_uses_canonical_values() -> None:
    state: dict[str, object] = {}
    app.initialize_setup_state_once(state=state)
    state[app.SETUP_STATE_KEY].update(
        {
            "survey_period": "06.00 - 19.00",
            "peak_mode": "rolling_60min",
            "am_peak_window_start": time(6, 0),
            "am_peak_window_end": time(10, 0),
            "pm_peak_window_start": time(15, 0),
            "pm_peak_window_end": time(19, 0),
        }
    )
    for field in app.ANALYZE_SETUP_FIELDS:
        state.pop(app.SETUP_FIELD_WIDGET_KEYS[field], None)

    restored = app._rehydrate_analyze_setup_widgets(state)

    assert restored["survey_period"] == "06.00 - 19.00"
    assert restored["peak_mode"] == "rolling_60min"
    assert restored["am_peak_window_start"] == time(6, 0)
    assert restored["am_peak_window_end"] == time(10, 0)
    assert restored["pm_peak_window_start"] == time(15, 0)
    assert restored["pm_peak_window_end"] == time(19, 0)
    assert state[app.SETUP_FIELD_WIDGET_KEYS["am_peak_window_start"]] == time(6, 0)


def test_single_analyze_setup_rehydration_preserves_revisions_readiness_and_artifacts() -> None:
    _seed_single_state()
    st.session_state["tmc_processed"]["result"] = SimpleNamespace(
        peaks=pd.DataFrame(
            [
                {"period": "AM", "peak_start": time(8, 0), "peak_end": time(9, 0), "hourly_pcu": 100},
                {"period": "PM", "peak_start": time(17, 0), "peak_end": time(18, 0), "hourly_pcu": 120},
            ]
        )
    )
    app._sync_single_workflow_from_state(
        source_bytes=b"source-a",
        source_file_name="demo.xlsx",
        export_mode=app.SAFE_PNG_EXPORT_MODE,
    )
    previous = app._workflow_state_for_mode(app.WORKFLOW_SINGLE_MODE)
    processed = st.session_state["tmc_processed"]
    output = st.session_state["tmc_output"]
    assert previous is not None
    assert previous.readiness == WorkflowReadiness(source=True, mapping=True, analysis=True, review=True, export=True)

    for field in app.ANALYZE_SETUP_FIELDS:
        st.session_state.pop(app.SETUP_FIELD_WIDGET_KEYS[field], None)
    app._rehydrate_analyze_setup_widgets()
    transition = app._sync_single_workflow_from_state(
        source_bytes=b"source-a",
        source_file_name="demo.xlsx",
        export_mode=app.SAFE_PNG_EXPORT_MODE,
    )
    current = app._workflow_state_for_mode(app.WORKFLOW_SINGLE_MODE)

    assert transition.changed_fields == ()
    assert current is not None
    assert current.revisions == previous.revisions
    assert current.readiness == previous.readiness
    assert st.session_state["tmc_processed"] is processed
    assert st.session_state["tmc_output"] is output


def test_batch_analyze_setup_rehydration_preserves_revisions_readiness_and_artifacts() -> None:
    upload, preset = _seed_batch_state()
    previous = app._workflow_state_for_mode(app.WORKFLOW_BATCH_MODE)
    analysis = st.session_state["tmc_batch_analysis_result"]
    output = st.session_state["tmc_batch_export_result"]
    assert previous is not None

    for field in app.ANALYZE_SETUP_FIELDS:
        st.session_state.pop(app.SETUP_FIELD_WIDGET_KEYS[field], None)
    app._rehydrate_analyze_setup_widgets()
    transition = app._sync_batch_workflow_from_state(
        batch_uploads=[upload],
        mapping_preset=preset,
        movement_code_scheme="from_to",
        metadata_rows=st.session_state["tmc_batch_file_metadata_table"],
        export_mode=app.BATCH_SAFE_PNG_EXPORT_LABEL,
    )
    current = app._workflow_state_for_mode(app.WORKFLOW_BATCH_MODE)

    assert transition.changed_fields == ()
    assert current is not None
    assert current.revisions == previous.revisions
    assert current.readiness == previous.readiness
    assert st.session_state["tmc_batch_analysis_result"] is analysis
    assert st.session_state["tmc_batch_export_result"] is output


def test_peak_search_change_stales_analysis_but_peak_decision_only_stales_export() -> None:
    _seed_single_state()
    st.session_state["peak_mode_select"] = "rolling_60min"
    st.session_state["am_peak_window_start_input"] = time(8, 0)
    peak_transition = _sync()

    assert peak_transition.analysis_invalidated is True
    assert st.session_state["tmc_pce_results_stale"] is True
    stored = app._workflow_state_for_mode(app.WORKFLOW_SINGLE_MODE)
    assert stored is not None
    assert stored.readiness == WorkflowReadiness(source=True, mapping=True, analysis=False, review=False, export=False)
    shell = _single_shell_state()
    assert shell["readiness"] == {"source": True, "mapping": True, "analysis": False, "review": False, "export": False}
    assert shell["steps"][2:] == ["active", "pending", "pending"]

    _seed_single_state()
    st.session_state["tmc_confirmed_am_peak_start"] = "09:00"
    st.session_state["tmc_confirmed_am_peak_end"] = "10:00"
    review_transition = _sync()

    assert review_transition.analysis_invalidated is False
    assert review_transition.review_updated is True
    assert review_transition.export_invalidated is True
    assert "tmc_processed" in st.session_state
    assert "tmc_output" not in st.session_state
    stored = app._workflow_state_for_mode(app.WORKFLOW_SINGLE_MODE)
    assert stored is not None
    assert stored.readiness == WorkflowReadiness(source=True, mapping=True, analysis=True, review=True, export=False)
    shell = _single_shell_state()
    assert shell["readiness"] == {"source": True, "mapping": True, "analysis": True, "review": True, "export": False}
    assert shell["steps"][2:] == ["completed", "ready", "active"]


def test_export_metadata_change_preserves_analysis_and_invalidates_artifact() -> None:
    _seed_single_state()
    st.session_state["tmc_title_input"] = "changed report title"

    transition = _sync()

    assert transition.analysis_invalidated is False
    assert transition.review_invalidated is False
    assert transition.export_invalidated is True
    assert "tmc_processed" in st.session_state
    assert "tmc_output" not in st.session_state
    stored = app._workflow_state_for_mode(app.WORKFLOW_SINGLE_MODE)
    assert stored is not None
    assert stored.readiness == WorkflowReadiness(source=True, mapping=True, analysis=True, review=True, export=False)
    shell = _single_shell_state()
    assert shell["readiness"] == {"source": True, "mapping": True, "analysis": True, "review": True, "export": False}
    assert shell["steps"][2:] == ["completed", "ready", "active"]


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
    item = SimpleNamespace(
        confirmed_AM_peak="08:00-09:00",
        confirmed_PM_peak="17:00-18:00",
        file_name="demo.xlsx",
        folder_name="file_01_demo",
        survey_date_text="",
        output_stem="demo",
        notes="",
        workbook_bytes=source_bytes,
    )
    st.session_state["tmc_batch_analysis_result"] = SimpleNamespace(
        successful_items=[item],
        items=[item],
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


def test_batch_peak_search_change_stales_analysis_and_export() -> None:
    upload, preset = _seed_batch_state()
    st.session_state["am_peak_window_start_input"] = time(8, 0)

    transition = app._sync_batch_workflow_from_state(
        batch_uploads=[upload],
        mapping_preset=preset,
        movement_code_scheme="from_to",
        metadata_rows=st.session_state["tmc_batch_file_metadata_table"],
        export_mode=app.BATCH_SAFE_PNG_EXPORT_LABEL,
    )

    stored = app._workflow_state_for_mode(app.WORKFLOW_BATCH_MODE)
    assert transition.analysis_invalidated is True
    assert st.session_state["tmc_batch_stale"] is True
    assert "tmc_batch_export_result" not in st.session_state
    assert stored is not None
    assert stored.readiness == WorkflowReadiness(source=True, mapping=True, analysis=False, review=False, export=False)


def test_batch_mapping_preset_change_updates_stored_readiness() -> None:
    upload, preset = _seed_batch_state()
    changed_preset = {
        **preset,
        "mapping_rows": [{"raw_sheet": "North", "output_movement_code": "NS"}],
    }

    transition = app._sync_batch_workflow_from_state(
        batch_uploads=[upload],
        mapping_preset=changed_preset,
        movement_code_scheme="from_to",
        metadata_rows=st.session_state["tmc_batch_file_metadata_table"],
        export_mode=app.BATCH_SAFE_PNG_EXPORT_LABEL,
    )

    stored = app._workflow_state_for_mode(app.WORKFLOW_BATCH_MODE)
    assert transition.analysis_invalidated is True
    assert st.session_state["tmc_batch_stale"] is True
    assert stored is not None
    assert stored.readiness.analysis is False
    assert stored.readiness.review is False
    assert stored.readiness.export is False
    shell = _batch_shell_state()
    assert shell["readiness"] == {"source": True, "mapping": True, "analysis": False, "review": False, "export": False}
    assert shell["steps"][2:] == ["active", "pending", "pending"]


def test_workflow_shell_status_uses_authoritative_readiness_in_canonical_order() -> None:
    _seed_single_state()
    stored = app._workflow_state_for_mode(app.WORKFLOW_SINGLE_MODE)
    assert stored is not None

    shell = _single_shell_state()

    assert [row[0] for row in shell["summary"]] == app.workflow_stages_for_mode(app.WORKFLOW_SINGLE_MODE)
    assert len(shell["steps"]) == len(app.workflow_stages_for_mode(app.WORKFLOW_SINGLE_MODE))
    assert shell["readiness"] == {
        "source": stored.readiness.source,
        "mapping": stored.readiness.mapping,
        "analysis": stored.readiness.analysis,
        "review": stored.readiness.review,
        "export": stored.readiness.export,
    }


def test_canonical_stage_navigation_does_not_invalidate_batch_workflow() -> None:
    upload, preset = _seed_batch_state()
    previous = app._workflow_state_for_mode(app.WORKFLOW_BATCH_MODE)
    assert previous is not None

    for stage in app.workflow_stages_for_mode(app.WORKFLOW_BATCH_MODE):
        app.set_active_tab(stage)

    transition = app._sync_batch_workflow_from_state(
        batch_uploads=[upload],
        mapping_preset=preset,
        movement_code_scheme="from_to",
        metadata_rows=st.session_state["tmc_batch_file_metadata_table"],
        export_mode=app.BATCH_SAFE_PNG_EXPORT_LABEL,
    )
    current = app._workflow_state_for_mode(app.WORKFLOW_BATCH_MODE)

    assert transition.changed_fields == ()
    assert current is not None
    assert current.revisions == previous.revisions
    assert st.session_state["tmc_batch_analysis_result"] is not None
    assert st.session_state["tmc_batch_export_result"] is not None


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
    stored = app._workflow_state_for_mode(app.WORKFLOW_BATCH_MODE)
    assert stored is not None
    assert stored.readiness == WorkflowReadiness(source=True, mapping=True, analysis=True, review=True, export=False)
    shell = _batch_shell_state()
    assert shell["readiness"] == {"source": True, "mapping": True, "analysis": True, "review": True, "export": False}
    assert shell["steps"][2:] == ["completed", "ready", "active"]


def test_batch_export_metadata_change_preserves_analysis_and_stales_export() -> None:
    upload, preset = _seed_batch_state()
    metadata_rows = [
        {"file_name": "demo.xlsx", "survey_date_text": "", "output_stem": "renamed", "notes": ""}
    ]
    st.session_state["tmc_batch_file_metadata_table"] = metadata_rows
    app._sync_batch_analysis_metadata_from_state()
    app._mark_batch_export_stale_now()

    transition = app._sync_batch_workflow_from_state(
        batch_uploads=[upload],
        mapping_preset=preset,
        movement_code_scheme="from_to",
        metadata_rows=metadata_rows,
        export_mode=app.BATCH_SAFE_PNG_EXPORT_LABEL,
    )

    assert transition.analysis_invalidated is False
    assert transition.export_invalidated is True
    assert transition.changed_fields == ("export_config",)
    assert st.session_state["tmc_batch_analysis_result"] is not None
    assert "tmc_batch_export_result" not in st.session_state
    stored = app._workflow_state_for_mode(app.WORKFLOW_BATCH_MODE)
    assert stored is not None
    assert stored.readiness.analysis is True
    assert stored.readiness.review is True
    assert stored.readiness.export is False
    shell = _batch_shell_state()
    assert shell["readiness"] == {"source": True, "mapping": True, "analysis": True, "review": True, "export": False}
    assert shell["steps"][2:] == ["completed", "ready", "active"]
