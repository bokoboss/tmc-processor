"""Data-stage rendering boundary."""

from __future__ import annotations

from tmc_processor.ui.workflow_context import WorkflowContext

from .batch import render_batch_data
from .single import render_single_data


def render_data_stage(*, context: WorkflowContext) -> None:
    if context.is_single_file_mode:
        render_single_data(context=context)
        return
    render_batch_data(context=context)
    return
