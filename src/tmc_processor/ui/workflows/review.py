"""Review-stage rendering boundary."""

from __future__ import annotations

from tmc_processor.ui.workflow_context import WorkflowContext

from .batch import render_batch_review
from .single import render_single_review
from tmc_processor.ui.components.peak import render_peak_card
from tmc_processor.ui.components.qc import render_qc_summary


def render_review_stage(*, context: WorkflowContext) -> None:
    if context.is_single_file_mode:
        render_single_review(context=context)
        return
    render_batch_review(context=context)
    return
