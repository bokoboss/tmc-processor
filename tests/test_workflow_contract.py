from __future__ import annotations

from tmc_processor.workflow_state import (
    WorkflowReadiness,
    WorkflowRevisions,
    WorkflowState,
    analysis_config_fingerprint,
    export_config_fingerprint,
    mapping_fingerprint,
    review_decision_fingerprint,
    source_fingerprint,
    transition_workflow,
)


def _revisions(**changes: str | None) -> WorkflowRevisions:
    base = WorkflowRevisions(
        source="source-a",
        mapping="mapping-a",
        analysis_config="analysis-a",
        analysis_result="result-a",
        review_decision="review-a",
        export_config="export-a",
    )
    return base.with_updates(**changes)


def test_same_semantic_source_and_configuration_does_not_invalidate() -> None:
    transition = transition_workflow(_revisions(), _revisions())

    assert transition.changed_fields == ()
    assert not transition.analysis_invalidated
    assert not transition.review_invalidated
    assert not transition.export_invalidated


def test_source_bytes_change_is_detected_even_when_filename_and_size_match() -> None:
    first = source_fingerprint(b"workbook-a")
    second = source_fingerprint(b"workbook-b")

    assert len(b"workbook-a") == len(b"workbook-b")
    assert first != second
    transition = transition_workflow(_revisions(source=first), _revisions(source=second))
    assert transition.analysis_invalidated
    assert transition.review_invalidated
    assert transition.export_invalidated


def test_mapping_semantic_content_change_invalidates_downstream_state() -> None:
    first = mapping_fingerprint([{"raw_sheet": "North", "movement_code": "NB"}])
    second = mapping_fingerprint([{"raw_sheet": "North", "movement_code": "NS"}])

    transition = transition_workflow(_revisions(mapping=first), _revisions(mapping=second))
    assert transition.analysis_invalidated
    assert transition.review_invalidated
    assert transition.export_invalidated


def test_mapping_key_order_and_editor_version_are_not_semantic_changes() -> None:
    first = mapping_fingerprint([{"raw_sheet": "North", "movement_code": "NB"}])
    second = mapping_fingerprint([{"movement_code": "NB", "raw_sheet": "North"}])

    assert first == second
    assert transition_workflow(_revisions(mapping=first), _revisions(mapping=second)).changed_fields == ()


def test_pce_factor_change_invalidates_analysis_review_and_export() -> None:
    first = analysis_config_fingerprint(pce_factors={"MC": 1.0}, peak_mode="fixed")
    second = analysis_config_fingerprint(pce_factors={"MC": 1.2}, peak_mode="fixed")
    transition = transition_workflow(_revisions(analysis_config=first), _revisions(analysis_config=second))

    assert transition.analysis_invalidated
    assert transition.review_invalidated
    assert transition.export_invalidated


def test_peak_search_mode_or_window_change_invalidates_analysis_review_and_export() -> None:
    first = analysis_config_fingerprint(
        pce_factors={"MC": 1.0},
        peak_mode="fixed",
        peak_windows={"AM": ("07:00", "12:00")},
    )
    second = analysis_config_fingerprint(
        pce_factors={"MC": 1.0},
        peak_mode="rolling",
        peak_windows={"AM": ("08:00", "12:00")},
    )
    transition = transition_workflow(_revisions(analysis_config=first), _revisions(analysis_config=second))

    assert transition.analysis_invalidated
    assert transition.review_invalidated
    assert transition.export_invalidated


def test_confirmed_peak_change_keeps_analysis_and_updates_review_but_invalidates_export() -> None:
    first = review_decision_fingerprint({"AM": "08:00-09:00", "PM": "17:00-18:00"})
    second = review_decision_fingerprint({"AM": "09:00-10:00", "PM": "17:00-18:00"})
    transition = transition_workflow(_revisions(review_decision=first), _revisions(review_decision=second))

    assert not transition.analysis_invalidated
    assert not transition.review_invalidated
    assert transition.review_updated
    assert transition.export_invalidated


def test_export_metadata_change_only_invalidates_export() -> None:
    first = export_config_fingerprint({"tmc_title": "A", "export_mode": "safe"})
    second = export_config_fingerprint({"tmc_title": "B", "export_mode": "safe"})
    transition = transition_workflow(_revisions(export_config=first), _revisions(export_config=second))

    assert not transition.analysis_invalidated
    assert not transition.review_invalidated
    assert transition.export_invalidated


def test_view_only_state_is_outside_semantic_revisions() -> None:
    previous = _revisions()
    current = _revisions()
    view_only_state = {"active_tab": "Review", "expanded_panel": "QC", "flash": "saved", "probe": 1}

    assert view_only_state
    assert transition_workflow(previous, current).changed_fields == ()


def test_analysis_result_revision_invalidates_review_and_export_not_analysis() -> None:
    transition = transition_workflow(_revisions(analysis_result="result-a"), _revisions(analysis_result="result-b"))

    assert not transition.analysis_invalidated
    assert transition.review_invalidated
    assert transition.export_invalidated


def test_workflow_state_exposes_mode_revisions_and_readiness_without_ui_dependencies() -> None:
    state = WorkflowState(
        mode="single",
        revisions=_revisions(),
        readiness=WorkflowReadiness(source=True, mapping=True, analysis=True, review=True, export=False),
    )

    assert state.mode == "single"
    assert state.revisions.source == "source-a"
    assert state.readiness.analysis is True
    assert state.readiness.export is False
