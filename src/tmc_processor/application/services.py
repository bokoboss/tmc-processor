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


def analyze_single(source: Any, mapping: Any, setup: dict[str, Any], **kwargs: Any) -> Any:
    """Run the existing Single domain pipeline."""

    return process_tmc(raw_sheets=source, mapping=mapping, setup=dict(setup), **kwargs)


def analyze_single_dry_run(source: Any, mapping: Any, setup: dict[str, Any], **kwargs: Any) -> Any:
    return process_tmc_dry_run_v2(raw_sheets=source, mapping=mapping, setup=dict(setup), **kwargs)


def analyze_batch(sources: Sequence[Any], **kwargs: Any) -> Any:
    return analyze_batch_files(list(sources), **kwargs)


def export_single_generated(**kwargs: Any) -> Any:
    return export_v2_generated_workbook(**kwargs)


def export_single_template_com(**kwargs: Any) -> Any:
    return export_v2_template_workbook_com(**kwargs)


def export_batch_reviewed(**kwargs: Any) -> Any:
    return generate_batch_zip_from_reviewed_peaks(**kwargs)
