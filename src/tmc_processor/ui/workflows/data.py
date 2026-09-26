"""Data-stage rendering boundary."""

from __future__ import annotations

from collections.abc import Mapping

import streamlit as st

from .batch import render_batch_stage
from .single import render_single_stage


def render_data_stage(*, context: Mapping[str, object]) -> None:
    if context.get("is_single_file_mode"):
        render_single_stage(context=context)
        return
    render_batch_stage(context=context)
    return
