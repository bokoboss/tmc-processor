"""Application orchestration services.

Services compose the established domain functions.  They do not implement
calculation, mapping, QC, Peak, or workbook-generation algorithms.
"""

from __future__ import annotations

from collections.abc import Sequence
from typing import Any

from tmc_processor.batch import analyze_batch_files, generate_batch_zip_from_reviewed_peaks
from tmc_processor.exporter import export_v2_generated_workbook, export_v2_template_workbook_com
from tmc_processor.pipeline import process_tmc, process_tmc_dry_run_v2


def analyze_single(source: Any = None, mapping: Any = None, setup: dict[str, Any] | None = None, **kwargs: Any) -> Any:
    """Run the existing Single domain pipeline."""

    source = kwargs.pop("raw_sheets", source)
    return process_tmc(raw_sheets=source, mapping=mapping, setup=dict(setup or {}), **kwargs)


def analyze_single_dry_run(source: Any = None, mapping: Any = None, setup: dict[str, Any] | None = None, **kwargs: Any) -> Any:
    source = kwargs.pop("raw_sheets", source)
    return process_tmc_dry_run_v2(raw_sheets=source, mapping=mapping, setup=dict(setup or {}), **kwargs)


def analyze_batch(sources: Sequence[Any], **kwargs: Any) -> Any:
    return analyze_batch_files(list(sources), **kwargs)


def export_single_generated(*args: Any, **kwargs: Any) -> Any:
    backend = kwargs.pop("_backend", export_v2_generated_workbook)
    return backend(*args, **kwargs)


def export_single_template_com(*args: Any, **kwargs: Any) -> Any:
    backend = kwargs.pop("_backend", export_v2_template_workbook_com)
    return backend(*args, **kwargs)


def export_batch_reviewed(*args: Any, **kwargs: Any) -> Any:
    backend = kwargs.pop("_backend", generate_batch_zip_from_reviewed_peaks)
    return backend(*args, **kwargs)
