"""Export-stage rendering boundary."""

from __future__ import annotations

from tmc_processor.ui.workflow_context import WorkflowContext

from .batch import render_batch_export
from .single import render_single_export
from tmc_processor.ui.components.export import render_export_backend


def render_export_stage(*, context: WorkflowContext) -> None:
    if context.is_single_file_mode:
        render_single_export(context=context)
        return
    render_batch_export(context=context)
    return
