from __future__ import annotations

from types import SimpleNamespace

import pandas as pd
import streamlit as st

import app
from tmc_processor.batch import (
    BATCH_SUMMARY_COLUMNS,
    BatchAnalysisItem,
    BatchAnalysisResult,
    bulk_accept_clean_files,
    batch_review_state,
    exclude_batch_item,
    eligible_clean_batch_items,
    reviewed_peak_values_complete,
    restore_batch_item,
)
from tmc_processor.exporter import (
    BATCH_SAFE_PNG_EXPORT_MODE,
    EXCEL_TEMPLATE_EXPORT_MODE,
    SAFE_PNG_EXPORT_MODE,
    STANDARD_REPORT_EXPORT_MODE,
    standard_report_export_decision,
    _export_metadata_frame,
)
from tmc_processor.peaks import PEAK_SELECTION_USER_CONFIRMED_BATCH


def _item(name: str, **changes: object) -> BatchAnalysisItem:
    values: dict[str, object] = {
        "file_name": name,
        "survey_date_text": "2026-02-03",
        "output_stem": name.removesuffix(".xlsx"),
        "folder_name": f"file_{name}",
        "status": "success",
        "suggested_AM_peak": "07:00-08:00",
        "suggested_PM_peak": "17:00-18:00",
    }
    values.update(changes)
    return BatchAnalysisItem(**values)


def test_standard_report_prefers_native_template_when_available_and_compatible() -> None:
    decision = standard_report_export_decision(
        excel_com_available=True,
        template_compatible=True,
    )

    assert decision.requested_mode == STANDARD_REPORT_EXPORT_MODE
    assert decision.backend_mode == EXCEL_TEMPLATE_EXPORT_MODE
    assert decision.use_template_report_layout is True
    assert decision.use_excel_com_native_charts is True
    assert decision.fallback_notice == ""


def test_standard_report_selects_safe_png_with_clear_fallback_when_com_is_unavailable() -> None:
    decision = standard_report_export_decision(
        excel_com_available=False,
        template_compatible=True,
        availability_detail="Excel COM is unavailable",
    )

    assert decision.backend_mode == SAFE_PNG_EXPORT_MODE
    assert decision.use_template_report_layout is False
    assert decision.use_excel_com_native_charts is False
    assert "Safe PNG" in decision.fallback_notice
    assert "Excel COM is unavailable" in decision.fallback_notice


def test_standard_report_does_not_attempt_an_unsupported_template_path() -> None:
    decision = standard_report_export_decision(
        excel_com_available=True,
        template_compatible=False,
        availability_detail="validated template map is not compatible",
    )

    assert decision.backend_mode == SAFE_PNG_EXPORT_MODE
    assert decision.use_template_report_layout is False
    assert decision.use_excel_com_native_charts is False
    assert "not compatible" in decision.fallback_notice


def test_export_metadata_records_requested_used_and_fallback_backend() -> None:
    metadata = _export_metadata_frame(
        {},
        export_mode=SAFE_PNG_EXPORT_MODE,
        export_mode_requested=STANDARD_REPORT_EXPORT_MODE,
        export_mode_used=SAFE_PNG_EXPORT_MODE,
        fallback_notice="COM unavailable",
    )
    values = dict(zip(metadata["field"], metadata["value"]))

    assert values["export_mode_requested"] == STANDARD_REPORT_EXPORT_MODE
    assert values["export_mode_used"] == SAFE_PNG_EXPORT_MODE
    assert values["export_fallback_notice"] == "COM unavailable"


def test_clean_batch_eligibility_ignores_qc_info_but_excludes_warnings_and_errors() -> None:
    clean = _item("clean.xlsx", QC_info=2)
    warning = _item("warning.xlsx", QC_warnings=1)
    error = _item("error.xlsx", QC_errors=1)
    no_suggestion = _item("missing.xlsx", suggested_PM_peak="")
    confirmed = _item("confirmed.xlsx", confirmed_AM_peak="07:00-08:00", confirmed_PM_peak="17:00-18:00")

    eligible = eligible_clean_batch_items(
        BatchAnalysisResult(items=[clean, warning, error, no_suggestion, confirmed])
    )

    assert [item.file_name for item in eligible] == ["clean.xlsx"]


def test_bulk_acceptance_uses_batch_confirmation_provenance() -> None:
    clean = _item("clean.xlsx")
    warning = _item("warning.xlsx", QC_warnings=1)
    analysis = BatchAnalysisResult(items=[clean, warning])

    accepted = bulk_accept_clean_files(analysis)

    assert [item.file_name for item in accepted] == ["clean.xlsx"]
    assert clean.confirmed_AM_peak == clean.suggested_AM_peak
    assert clean.confirmed_PM_peak == clean.suggested_PM_peak
    assert clean.peak_selection_source == PEAK_SELECTION_USER_CONFIRMED_BATCH
    assert warning.confirmed_AM_peak == ""
    assert batch_review_state(clean) == "confirmed"
    assert batch_review_state(warning) == "needs_review"


def test_excluded_successful_file_satisfies_review_without_being_exportable() -> None:
    confirmed = _item("confirmed.xlsx", confirmed_AM_peak="07:00-08:00", confirmed_PM_peak="17:00-18:00")
    excluded = _item("excluded.xlsx")
    analysis = BatchAnalysisResult(items=[confirmed, excluded])

    exclude_batch_item(excluded, reason="Operator removed duplicate upload")

    assert batch_review_state(excluded) == "excluded"
    assert reviewed_peak_values_complete(analysis) is True
    assert excluded.exclusion_reason == "Operator removed duplicate upload"

    restore_batch_item(excluded)

    assert batch_review_state(excluded) == "needs_review"
    assert reviewed_peak_values_complete(analysis) is False


def test_batch_summary_contract_exposes_review_disposition_columns() -> None:
    assert "review_state" in BATCH_SUMMARY_COLUMNS
    assert "disposition_reason" in BATCH_SUMMARY_COLUMNS


def test_exclusion_invalidates_export_without_invalidating_analysis() -> None:
    st.session_state.clear()
    app.initialize_setup_state_once()
    item = _item("demo.xlsx")
    st.session_state["tmc_batch_analysis_result"] = BatchAnalysisResult(items=[item])
    st.session_state["tmc_batch_file_metadata_table"] = [
        {"file_name": "demo.xlsx", "survey_date_text": "", "output_stem": "demo", "notes": ""}
    ]
    upload = SimpleNamespace(name="demo.xlsx", getvalue=lambda: b"demo")
    app._sync_batch_workflow_from_state(
        batch_uploads=[upload],
        mapping_preset={},
        movement_code_scheme="from_to",
        metadata_rows=st.session_state["tmc_batch_file_metadata_table"],
        export_mode=app.SAFE_PNG_EXPORT_MODE,
    )
    st.session_state["tmc_batch_export_result"] = object()
    st.session_state["tmc_batch_export_stale"] = False

    assert app._exclude_batch_file(item.folder_name, "duplicate") is True
    transition = app._sync_batch_workflow_from_state(
        batch_uploads=[upload],
        mapping_preset={},
        movement_code_scheme="from_to",
        metadata_rows=st.session_state["tmc_batch_file_metadata_table"],
        export_mode=app.SAFE_PNG_EXPORT_MODE,
    )

    assert transition.analysis_invalidated is False
    assert transition.export_invalidated is True
    assert st.session_state["tmc_batch_analysis_result"] is not None
