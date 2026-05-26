from __future__ import annotations

from pathlib import Path

import pandas as pd
import pytest

from tmc_processor.exporter import export_workbook
from tmc_processor.importer import load_detected_sheets
from tmc_processor.mapping import apply_saved_mapping_to_sheets, read_mapping_excel_with_metadata, validate_mapping_scheme
from tmc_processor.mapping_preset import apply_mapping_preset_to_detected_sheets, load_mapping_preset
from tmc_processor.movement_scheme import APPROACH_MOVEMENT_CODES, MOVEMENT_SCHEME_V2, is_approach_movement_code
from tmc_processor.pipeline import process_tmc, process_tmc_dry_run_v2
from tmc_processor.batch import batch_processing_block_reason


ROOT = Path(__file__).resolve().parents[1]
DEMO_DIR = ROOT / "samples" / "demo"
RAW_WORKBOOK = DEMO_DIR / "DEMO_TMC1_FourLeg.xlsx"
V2_PRESET = DEMO_DIR / "DEMO_TMC1_FourLeg_approach_v2.mapping.json"
V2_MAPPING_XLSX = DEMO_DIR / "DEMO_TMC1_FourLeg_approach_v2_mapping.xlsx"


def _setup() -> dict[str, object]:
    return {
        "project_name": "Synthetic Demo Project",
        "tmc_id": "DEMO-TMC1",
        "tmc_name": "Synthetic Four-Leg Demo Intersection",
        "survey_date": "2026-01-01",
        "movement_code_scheme": MOVEMENT_SCHEME_V2,
    }


def _raw_sheets() -> dict[str, pd.DataFrame]:
    return load_detected_sheets(RAW_WORKBOOK)


def _v2_preset_mapping(raw_sheets: dict[str, pd.DataFrame]) -> pd.DataFrame:
    preset = load_mapping_preset(V2_PRESET.read_bytes()).preset
    return apply_mapping_preset_to_detected_sheets(preset, list(raw_sheets)).mapping


def _v2_excel_mapping(raw_sheets: dict[str, pd.DataFrame]) -> pd.DataFrame:
    loaded = read_mapping_excel_with_metadata(V2_MAPPING_XLSX)
    assert loaded.movement_code_scheme == MOVEMENT_SCHEME_V2
    return apply_saved_mapping_to_sheets(list(raw_sheets), loaded.mapping)


def _dry_run_with_preset():
    raw_sheets = _raw_sheets()
    mapping = _v2_preset_mapping(raw_sheets)
    return process_tmc_dry_run_v2(
        raw_sheets=raw_sheets,
        mapping=mapping,
        setup=_setup(),
        detected_sheets=list(raw_sheets),
    )


def test_v2_demo_mapping_preset_loads_and_validates_for_dry_run() -> None:
    raw_sheets = _raw_sheets()
    mapping = _v2_preset_mapping(raw_sheets)

    assert validate_mapping_scheme(mapping, MOVEMENT_SCHEME_V2) == []
    assert mapping["movement_code"].tolist() == ["NT", "NT", "WT", "ET"]
    assert mapping["approach_direction"].tolist() == ["N", "N", "W", "E"]
    assert mapping["movement_type"].tolist() == ["T", "T", "T", "T"]


def test_v2_mapping_excel_loads_and_dry_runs() -> None:
    raw_sheets = _raw_sheets()
    mapping = _v2_excel_mapping(raw_sheets)
    result = process_tmc_dry_run_v2(
        raw_sheets=raw_sheets,
        mapping=mapping,
        setup=_setup(),
        detected_sheets=list(raw_sheets),
    )

    assert not result.normalized.empty
    assert set(result.normalized["movement_code"]) == {"NT", "WT", "ET"}
    assert result.workbook_bytes == b""


def test_v2_dry_run_normalizes_approach_movement_codes_and_metadata() -> None:
    result = _dry_run_with_preset()

    assert not result.normalized.empty
    assert set(result.normalized["movement_code"]) == {"NT", "WT", "ET"}
    assert set(result.normalized["output_movement_code"]) == {"NT", "WT", "ET"}
    assert set(result.normalized["movement_code_scheme"]) == {MOVEMENT_SCHEME_V2}
    assert set(result.normalized["approach_direction"]) == {"N", "W", "E"}
    assert set(result.normalized["movement_type"]) == {"T"}
    assert all(is_approach_movement_code(code) for code in result.normalized["movement_code"].unique())


def test_v2_hourly_movement_summary_uses_approach_movement_order() -> None:
    result = _dry_run_with_preset()

    assert list(result.hourly_movement_pcu.columns[1:-1]) == APPROACH_MOVEMENT_CODES
    assert result.hourly_movement_pcu.columns[-1] == "Total"
    assert set(result.hourly_movement_pcu.columns[1:-1]) == set(APPROACH_MOVEMENT_CODES)


def test_v2_movement_codes_in_outputs_are_valid_approach_movement() -> None:
    result = _dry_run_with_preset()

    normalized_codes = set(result.normalized["movement_code"].dropna().astype(str))
    movement_codes = set(result.movement["movement_code"].dropna().astype(str))
    summary_codes = set(result.hourly_movement_pcu.columns[1:-1])

    assert normalized_codes
    assert normalized_codes <= set(APPROACH_MOVEMENT_CODES)
    assert movement_codes <= set(APPROACH_MOVEMENT_CODES)
    assert all(is_approach_movement_code(code) for code in summary_codes)


def test_v2_peak_calculation_returns_am_pm_without_v1_movement_assumptions() -> None:
    result = _dry_run_with_preset()

    assert set(result.peaks["period"].astype(str)) == {"AM", "PM"}
    assert result.peaks["hourly_pcu"].gt(0).all()
    assert set(result.normalized["movement_code"]).isdisjoint({"NS", "WE", "EW"})


def test_v2_dry_run_does_not_enable_final_excel_export() -> None:
    result = _dry_run_with_preset()

    assert result.workbook_bytes == b""
    with pytest.raises(ValueError, match="v2 export/report generation is not supported"):
        export_workbook(
            _setup(),
            pd.DataFrame(),
            result.normalized,
            result.qc,
            result.hourly,
            result.movement,
            result.vehicle,
            result.peaks,
        )

    with pytest.raises(ValueError, match="approach_movement v2"):
        process_tmc(
            raw_sheets=_raw_sheets(),
            mapping=_v2_preset_mapping(_raw_sheets()),
            setup=_setup(),
            detected_sheets=list(_raw_sheets()),
            generate_workbook=False,
        )


def test_mixed_v1_v2_mapping_codes_are_rejected() -> None:
    mixed = pd.DataFrame(
        [
            {"raw_sheet": "A", "movement_code": "NS", "include_in_report": True},
            {"raw_sheet": "B", "movement_code": "NT", "include_in_report": True},
        ]
    )

    assert validate_mapping_scheme(mixed, MOVEMENT_SCHEME_V2) == [
        "Mapping contains mixed from_to and approach_movement movement codes.",
        "Row 1 has invalid approach_movement output_movement_code 'NS'.",
    ]


def test_batch_v2_dry_run_is_explicitly_unsupported() -> None:
    assert "approach_movement v2" in batch_processing_block_reason(MOVEMENT_SCHEME_V2)
