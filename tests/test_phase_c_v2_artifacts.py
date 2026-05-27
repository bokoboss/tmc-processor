from __future__ import annotations

import json
from pathlib import Path

import pandas as pd

from tmc_processor.mapping import read_mapping_excel
from tmc_processor.mapping_preset import load_mapping_preset
from tmc_processor.movement_scheme import (
    APPROACH_MOVEMENT_CODES,
    MOVEMENT_SCHEME_V2,
    is_approach_movement_code,
)


ROOT = Path(__file__).resolve().parents[1]
V2_PRESET_PATH = ROOT / "samples/demo/DEMO_TMC1_FourLeg_approach_v2.mapping.json"
V2_MAPPING_XLSX_PATH = ROOT / "samples/demo/DEMO_TMC1_FourLeg_approach_v2_mapping.xlsx"
V2_TEMPLATE_MAP_PATH = ROOT / "templates/four_leg_tmc_report_template_approach_v2_map.json"
V1_PRESET_PATH = ROOT / "samples/demo/DEMO_TMC1_FourLeg.mapping.json"


def _json(path: Path) -> dict[str, object]:
    return json.loads(path.read_text(encoding="utf-8"))


def test_v2_mapping_preset_declares_approach_movement_scheme() -> None:
    preset = _json(V2_PRESET_PATH)

    assert preset["movement_code_scheme"] == MOVEMENT_SCHEME_V2
    assert preset["template_version"] == "four_leg_approach_movement_v2"
    assert preset["movement_code_order"] == APPROACH_MOVEMENT_CODES


def test_v2_mapping_preset_codes_are_valid_approach_movement_codes() -> None:
    preset = _json(V2_PRESET_PATH)
    rows = preset["mapping_rows"]

    assert isinstance(rows, list)
    assert rows
    output_codes = [str(row["output_movement_code"]) for row in rows if isinstance(row, dict)]

    assert output_codes == ["NT", "NT", "WT", "ET"]
    assert all(is_approach_movement_code(code) for code in output_codes)
    assert set(output_codes).issubset(APPROACH_MOVEMENT_CODES)


def test_v2_mapping_preset_remains_loadable_by_current_safe_loader() -> None:
    loaded = load_mapping_preset(V2_PRESET_PATH.read_bytes())

    assert loaded.warnings == ()
    assert loaded.preset["movement_code_scheme"] == MOVEMENT_SCHEME_V2
    assert [row["output_movement_code"] for row in loaded.preset["mapping_rows"]] == ["NT", "NT", "WT", "ET"]


def test_v2_mapping_excel_keeps_compatible_mapping_sheet_and_metadata() -> None:
    workbook = pd.ExcelFile(V2_MAPPING_XLSX_PATH)

    assert workbook.sheet_names == ["Mapping", "Metadata", "Movement_Code_Reference"]

    raw_mapping = pd.read_excel(workbook, sheet_name="Mapping")
    assert {"output_movement_code", "approach_direction", "movement_type"}.issubset(raw_mapping.columns)
    assert raw_mapping["output_movement_code"].tolist() == ["NT", "NT", "WT", "ET"]

    cleaned = read_mapping_excel(V2_MAPPING_XLSX_PATH)
    assert cleaned["movement_code"].tolist() == ["NT", "NT", "WT", "ET"]

    metadata = pd.read_excel(workbook, sheet_name="Metadata", header=None, index_col=0)[1].to_dict()
    assert metadata["movement_code_scheme"] == MOVEMENT_SCHEME_V2


def test_v2_template_map_references_expected_approach_movement_code_set() -> None:
    template_map = _json(V2_TEMPLATE_MAP_PATH)

    assert template_map["movement_code_scheme"] == MOVEMENT_SCHEME_V2
    assert template_map["template_version"] == "four_leg_approach_movement_v2"
    assert template_map["movement_code_order"] == APPROACH_MOVEMENT_CODES

    hourly_columns = template_map["hourly_movement_table"]["columns"]
    assert [code for code in hourly_columns if code not in {"time", "Total"}] == APPROACH_MOVEMENT_CODES

    diagram_movements = template_map["movement_diagram_cells"]["diagram_movements"]
    assert set(diagram_movements) == set(APPROACH_MOVEMENT_CODES)
    assert template_map["movement_diagram_cells"]["approach_tables"]["north"]["movement_codes"] == [
        "NL",
        "NT",
        "NR",
        "NU",
    ]


def test_existing_v1_demo_preset_remains_from_to() -> None:
    preset = _json(V1_PRESET_PATH)

    assert preset["movement_code_scheme"] == "from_to"
    assert [row["output_movement_code"] for row in preset["mapping_rows"]] == ["NS", "NS", "WE", "EW"]
