"""Pure workflow composition and transition orchestration."""

from __future__ import annotations

from dataclasses import dataclass
from collections.abc import MutableMapping, Sequence
from typing import Any

from tmc_processor.workflow_state import (
    WorkflowReadiness,
    WorkflowRevisions,
    WorkflowState,
    WorkflowTransition,
    analysis_config_fingerprint,
    analysis_result_fingerprint,
    export_config_fingerprint,
    mapping_fingerprint,
    readiness_after_transition,
    review_decision_fingerprint,
    semantic_fingerprint,
    source_fingerprint,
    transition_workflow,
)

CANONICAL_WORKFLOW_STAGES = ("Data", "Mapping", "Analyze", "Review", "Export")
WORKFLOW_SINGLE_MODE = "single"
WORKFLOW_BATCH_MODE = "batch"


def canonical_workflow_stage(value: object) -> str:
    candidate = str(value or "")
    return candidate if candidate in CANONICAL_WORKFLOW_STAGES else CANONICAL_WORKFLOW_STAGES[0]


def workflow_stages_for_mode(mode: str) -> list[str]:
    if mode not in {WORKFLOW_SINGLE_MODE, WORKFLOW_BATCH_MODE}:
        raise ValueError(f"Unsupported workflow mode: {mode}")
    return list(CANONICAL_WORKFLOW_STAGES)


def _peak_payload(peaks: dict[str, object] | None) -> dict[str, str] | None:
    if not peaks:
        return None
    required = ("am_peak_start", "am_peak_end", "pm_peak_start", "pm_peak_end")
    if not all(str(peaks.get(key, "")).strip() for key in required):
        return None
    return {key: str(peaks[key]) for key in required}


def _analysis_revision(revisions: WorkflowRevisions, marker: object) -> str | None:
    return None if marker is None else analysis_result_fingerprint(
        {"source": revisions.source, "mapping": revisions.mapping, "analysis_config": revisions.analysis_config, "marker": marker}
    )


def single_workflow_revisions(
    *,
    source_bytes: bytes | None = None,
    source_revision: object = None,
    mapping: object,
    pce_factors: dict[str, float],
    peak_mode: str | None,
    peak_windows: dict[str, tuple[object, object]],
    movement_code_scheme: str,
    export_config: object,
    confirmed_peaks: dict[str, object] | None = None,
    analysis_present: bool = False,
) -> WorkflowRevisions:
    source = source_fingerprint(source_bytes) if source_revision is None else str(source_revision or "") or None
    revisions = WorkflowRevisions(
        source=source,
        mapping=mapping_fingerprint(mapping),
        analysis_config=analysis_config_fingerprint(
            pce_factors=pce_factors,
            peak_mode=peak_mode,
            peak_windows=peak_windows,
            movement_code_scheme=movement_code_scheme,
        ),
        review_decision=review_decision_fingerprint(_peak_payload(confirmed_peaks)),
        export_config=export_config_fingerprint(export_config),
    )
    return revisions.with_updates(analysis_result=_analysis_revision(revisions, "single")) if analysis_present else revisions


def batch_workflow_revisions(
    *,
    uploads: Sequence[object] | None,
    mapping_preset: dict[str, object] | None,
    pce_factors: dict[str, float],
    peak_mode: str | None,
    peak_windows: dict[str, tuple[object, object]],
    movement_code_scheme: str,
    metadata_rows: list[dict[str, object]],
    export_config: object,
    confirmed_peaks: object = None,
    dispositions: object = None,
    analysis_present: bool = False,
) -> WorkflowRevisions:
    content = tuple(source_fingerprint(getattr(file, "getvalue")()) for file in uploads or [])
    source = semantic_fingerprint(content) if content else None
    mapping = semantic_fingerprint({"movement_code_scheme": movement_code_scheme, "mapping_rows": (mapping_preset or {}).get("mapping_rows", [])})
    review_payload: dict[str, dict[str, str]] = {}
    for folder, values in (confirmed_peaks or {}).items() if isinstance(confirmed_peaks, dict) else []:
        values = values or {}
        if str(values.get("AM") or "") and str(values.get("PM") or ""):
            review_payload[str(folder)] = {"AM": str(values["AM"]), "PM": str(values["PM"])}
    for folder, values in (dispositions or {}).items() if isinstance(dispositions, dict) else []:
        if str((values or {}).get("disposition") or "").casefold() == "excluded":
            review_payload[str(folder)] = {"disposition": "excluded", "reason": str((values or {}).get("reason") or "")}
    revisions = WorkflowRevisions(
        source=source,
        mapping=mapping,
        analysis_config=analysis_config_fingerprint(
            pce_factors=pce_factors,
            peak_mode=peak_mode,
            peak_windows=peak_windows,
            movement_code_scheme=movement_code_scheme,
        ),
        review_decision=review_decision_fingerprint(review_payload),
        export_config=export_config_fingerprint(export_config),
    )
    return revisions.with_updates(analysis_result=_analysis_revision(revisions, "batch")) if analysis_present else revisions


def apply_transition_to_state(state: MutableMapping[str, object], transition: WorkflowTransition) -> None:
    """Apply only readiness/invalidation effects to a plain state mapping."""

    if not transition.engineering_state_changed:
        return
    if transition.analysis_invalidated:
        state["analysis_stale"] = True
        state["review_ready"] = False
        state["export_ready"] = False
    elif transition.review_invalidated:
        state["review_ready"] = False
        state["export_ready"] = False
    elif transition.export_invalidated:
        state["export_ready"] = False


@dataclass(frozen=True)
class WorkflowApplicationResult:
    state: WorkflowState
    transition: WorkflowTransition


def synchronize_workflow(
    previous: WorkflowState | None,
    *,
    mode: str,
    revisions: WorkflowRevisions,
    readiness: WorkflowReadiness,
    state: MutableMapping[str, object] | None = None,
) -> WorkflowApplicationResult:
    transition = transition_workflow(previous.revisions if previous else None, revisions)
    if state is not None:
        apply_transition_to_state(state, transition)
    if transition.analysis_invalidated:
        revisions = revisions.with_updates(analysis_result=None)
    current = WorkflowState(mode=mode, revisions=revisions, readiness=readiness_after_transition(readiness, transition))
    return WorkflowApplicationResult(state=current, transition=transition)
