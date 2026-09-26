"""Mapping-stage rendering boundary."""

from __future__ import annotations

from tmc_processor.ui.workflow_context import WorkflowContext

from .batch import render_batch_mapping
from .single import render_single_mapping
from tmc_processor.ui.components.mapping import render_mapping_summary


def render_mapping_stage(*, context: WorkflowContext) -> None:
    if context.is_single_file_mode:
        render_single_mapping(context=context)
        return
    render_batch_mapping(context=context)
    return
