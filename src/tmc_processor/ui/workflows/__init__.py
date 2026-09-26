"""Canonical workflow-stage render entrypoints.

The shell selects a canonical stage once and hands the stage context to this
registry.  Individual modules own the user-facing rendering for that stage;
the context keeps mode-specific orchestration explicit at the UI boundary.
"""

from .analyze import render_analyze_stage
from .data import render_data_stage
from .export import render_export_stage
from .mapping import render_mapping_stage
from .review import render_review_stage

WORKFLOW_STAGE_RENDERERS = {
    "Data": render_data_stage,
    "Mapping": render_mapping_stage,
    "Analyze": render_analyze_stage,
    "Review": render_review_stage,
    "Export": render_export_stage,
}


def render_workflow_stage(stage: str, *, context: dict[str, object]) -> object:
    """Render one canonical stage through its owning workflow module."""

    try:
        renderer = WORKFLOW_STAGE_RENDERERS[stage]
    except KeyError as exc:
        raise ValueError(f"Unknown workflow stage: {stage}") from exc
    return renderer(context=context)

__all__ = [
    "render_analyze_stage",
    "render_data_stage",
    "render_export_stage",
    "render_mapping_stage",
    "render_review_stage",
    "WORKFLOW_STAGE_RENDERERS",
    "render_workflow_stage",
]
