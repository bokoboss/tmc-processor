"""Analyze-stage rendering boundary."""

from __future__ import annotations

from tmc_processor.ui.workflow_context import WorkflowContext

from .batch import render_batch_analyze
from .single import render_single_analyze


def render_analyze_stage(*, context: WorkflowContext) -> None:
    if context.is_single_file_mode:
        render_single_analyze(context=context)
        return
    render_batch_analyze(context=context)
    return
