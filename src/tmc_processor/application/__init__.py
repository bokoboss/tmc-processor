"""Application-layer contracts for the Streamlit application shell.

The application package contains state adapters, workflow orchestration, and
service composition.  It deliberately has no dependency on Streamlit or the
UI package.
"""

from .state import (
    BATCH_CONFIRMED_PEAKS_KEY,
    BATCH_DISPOSITION_KEY,
    BATCH_DRAFT_PEAKS_KEY,
    SINGLE_CONFIRMED_PEAKS_KEY,
    clear_batch_review_state,
    clear_single_review_state,
    confirm_batch_peak,
    confirm_single_peaks,
    get_batch_dispositions,
    get_confirmed_peaks,
    get_draft_peaks,
    set_batch_disposition,
    set_draft_peaks,
)
from .workflow import (
    CANONICAL_WORKFLOW_STAGES,
    WorkflowApplicationResult,
    apply_transition_to_state,
    batch_workflow_revisions,
    canonical_workflow_stage,
    single_workflow_revisions,
    synchronize_workflow,
    workflow_stages_for_mode,
)

__all__ = [
    "BATCH_CONFIRMED_PEAKS_KEY",
    "BATCH_DISPOSITION_KEY",
    "BATCH_DRAFT_PEAKS_KEY",
    "CANONICAL_WORKFLOW_STAGES",
    "SINGLE_CONFIRMED_PEAKS_KEY",
    "WorkflowApplicationResult",
    "apply_transition_to_state",
    "batch_workflow_revisions",
    "canonical_workflow_stage",
    "clear_batch_review_state",
    "clear_single_review_state",
    "confirm_batch_peak",
    "confirm_single_peaks",
    "get_batch_dispositions",
    "get_confirmed_peaks",
    "get_draft_peaks",
    "set_batch_disposition",
    "set_draft_peaks",
    "single_workflow_revisions",
    "synchronize_workflow",
    "workflow_stages_for_mode",
]
