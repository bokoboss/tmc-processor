"""Mapping-stage rendering boundary."""

from __future__ import annotations

from collections.abc import Mapping

import streamlit as st

from .batch import render_batch_stage
from .single import render_single_stage
from tmc_processor.ui.components.mapping import render_mapping_summary


def render_mapping_stage(*, context: Mapping[str, object]) -> None:
    if context.get("is_single_file_mode"):
        render_single_stage(context=context)
        return
    render_batch_stage(context=context)
    return
