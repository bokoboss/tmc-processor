from __future__ import annotations

from io import BytesIO
import json
from pathlib import Path

import pandas as pd
import pytest

from tmc_processor.mapping import (
    APPROACH_MOVEMENT_PROCESSING_BLOCK_REASON,
    apply_saved_mapping_to_sheets,
    mapping_is_process_compatible,
    mapping_processing_block_reason,
    mapping_to_excel_bytes,
    read_mapping_excel,
    read_mapping_excel_with_metadata,
)
from tmc_processor.mapping_preset import (
    MAPPING_PRESET_TYPE,
    MappingPresetError,
    apply_mapping_preset_to_detected_sheets,
    build_mapping_preset,
    detect_mapping_preset_scheme,
    load_mapping_preset,
    mapping_preset_metadata,
    serialize_mapping_preset,
)
from tmc_processor.movement_scheme import MOVEMENT_SCHEME_V1, MOVEMENT_SCHEME_V2
from tmc_processor.pipeline import process_tmc
from tmc_processor.session import apply_session_to_state, build_project_session, session_from_json, session_to_json_bytes


ROOT = Path(__file__).resolve().parents[1]
DEMO_DIR = ROOT / "samples" / "demo"
V2_PRESET = DEMO_DIR / "DEMO_TMC1_FourLeg_approach_v2.mapping.json"
V2_MAPPING_XLSX = DEMO_DIR / "DEMO_TMC1_FourLeg_approach_v2_mapping.xlsx"


def _mapping() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "raw_sheet": "North mainline",
                "raw_direction": "North",
                "movement_code": "NS",
                "source_stream": "mainline",
                "raw_movement_label": "North through",
                "from_leg": "N",
                "to_leg": "S",
                "turn_type": "through",
                "facility_type": "at_grade",
                "include_in_peak": True,
                "include_in_report": True,
                "aggregation_method": "sum",
                "remark": "primary stream",
            },
            {
                "raw_sheet": "North frontage",
                "raw_direction": "North",
                "movement_code": "NS",
                "source_stream": "frontage",
                "raw_movement_label": "North frontage through",
                "from_leg": "N",
                "to_leg": "S",
                "turn_type": "through",
                "facility_type": "frontage",
                "include_in_peak": True,
                "include_in_report": True,
                "aggregation_method": "sum",
            },
            {
                "raw_sheet": "East",
                "raw_direction": "East",
                "movement_code": "EW",
                "source_stream": "mainline",
                "raw_movement_label": "East through",
                "from_leg": "E",
                "to_leg": "W",
                "turn_type": "through",
                "facility_type": "at_grade",
                "include_in_peak": False,
                "include_in_report": True,
                "aggregation_method": "sum",
            },
        ]
    )


def test_mapping_preset_serialization_contains_mapping_only_fields() -> None:
    preset = build_mapping_preset(_mapping(), preset_name="Demo preset", notes="Reusable mapping")
    loaded = load_mapping_preset(serialize_mapping_preset(preset)).preset

    assert loaded["preset_type"] == MAPPING_PRESET_TYPE
    assert loaded["preset_name"] == "Demo preset"
    assert loaded["notes"] == "Reusable mapping"
    assert "mapping_rows" in loaded
    assert "pce_factors" not in loaded
    assert "peaks" not in loaded
    assert "metadata" not in loaded
    assert loaded["mapping_rows"][0]["output_movement_code"] == "NS"


def test_old_v1_preset_without_scheme_defaults_to_from_to() -> None:
    preset = build_mapping_preset(_mapping())
    preset.pop("movement_code_scheme")

    loaded = load_mapping_preset(serialize_mapping_preset(preset)).preset

    assert loaded["movement_code_scheme"] == MOVEMENT_SCHEME_V1
    assert detect_mapping_preset_scheme(loaded) == MOVEMENT_SCHEME_V1


def test_mapping_preset_load_round_trips_stream_and_aggregation_fields() -> None:
    loaded = load_mapping_preset(serialize_mapping_preset(build_mapping_preset(_mapping()))).preset
    rows = loaded["mapping_rows"]

    assert rows[1]["source_stream"] == "frontage"
    assert rows[1]["aggregation_method"] == "sum"
    assert rows[1]["raw_movement_label"] == "North frontage through"
    assert rows[2]["include_in_peak"] is False


def test_v2_demo_preset_detects_and_validates_approach_movement() -> None:
    loaded = load_mapping_preset(V2_PRESET.read_bytes()).preset

    assert detect_mapping_preset_scheme(loaded) == MOVEMENT_SCHEME_V2
    assert mapping_preset_metadata(loaded).template_version == "four_leg_approach_movement_v2"
    assert loaded["movement_code_order"][0] == "NL"
    assert [row["output_movement_code"] for row in loaded["mapping_rows"]] == ["NT", "NT", "WT", "ET"]


def test_v2_preset_rejects_v1_or_unknown_codes() -> None:
    preset = json.loads(V2_PRESET.read_text(encoding="utf-8"))
    for invalid in ["NS", "WE", "XX"]:
        bad = dict(preset)
        bad["mapping_rows"] = [dict(preset["mapping_rows"][0], output_movement_code=invalid)]
        with pytest.raises(MappingPresetError, match="invalid approach_movement"):
            load_mapping_preset(json.dumps(bad).encode("utf-8"))


def test_v2_extra_approach_fields_survive_load_and_save() -> None:
    loaded = load_mapping_preset(V2_PRESET.read_bytes()).preset

    assert loaded["mapping_rows"][0]["approach_direction"] == "N"
    assert loaded["mapping_rows"][0]["movement_type"] == "T"

    frame = apply_mapping_preset_to_detected_sheets(loaded, [loaded["mapping_rows"][0]["raw_sheet"]]).mapping
    rebuilt = build_mapping_preset(frame, movement_code_scheme=MOVEMENT_SCHEME_V2)

    assert rebuilt["mapping_rows"][0]["approach_direction"] == "N"
    assert rebuilt["mapping_rows"][0]["movement_type"] == "T"


def test_apply_mapping_preset_matches_detected_sheets_and_preserves_duplicate_output_codes() -> None:
    loaded = load_mapping_preset(serialize_mapping_preset(build_mapping_preset(_mapping()))).preset
    result = apply_mapping_preset_to_detected_sheets(loaded, ["North mainline", "North frontage", "East"])

    assert result.matched_sheet_count == 3
    assert result.missing_detected_sheet_count == 0
    assert result.extra_preset_row_count == 0
    assert result.mapping["movement_code"].tolist().count("NS") == 2
    assert result.mapping.loc[result.mapping["raw_sheet"] == "North frontage", "source_stream"].iloc[0] == "frontage"


def test_missing_detected_sheets_remain_for_review() -> None:
    loaded = load_mapping_preset(serialize_mapping_preset(build_mapping_preset(_mapping()))).preset
    result = apply_mapping_preset_to_detected_sheets(loaded, ["North mainline", "South"])

    assert result.matched_sheet_count == 1
    assert result.missing_detected_sheet_count == 1
    assert "South" in result.missing_detected_sheets
    south = result.mapping[result.mapping["raw_sheet"] == "South"].iloc[0]
    assert south["movement_code"] == ""
    assert "Needs review" in south["remark"]


def test_extra_preset_rows_are_reported_without_entering_active_mapping() -> None:
    loaded = load_mapping_preset(serialize_mapping_preset(build_mapping_preset(_mapping()))).preset
    result = apply_mapping_preset_to_detected_sheets(loaded, ["North mainline"])

    assert result.extra_preset_row_count == 2
    assert set(result.extra_preset_sheets) == {"North frontage", "East"}
    assert result.mapping["raw_sheet"].tolist() == ["North mainline"]


def test_project_session_still_round_trips_mapping() -> None:
    session = build_project_session(mapping=_mapping(), detected_sheet_names=["North mainline"])
    loaded = session_from_json(session_to_json_bytes(session))
    state: dict[str, object] = {}

    apply_session_to_state(loaded.session, state)

    assert state["mapping_table"][0]["raw_sheet"] == "North mainline"
    assert state["mapping_editor_version"] == 1
    assert state["tmc_mapping_code_scheme"] == MOVEMENT_SCHEME_V1


def test_project_session_preserves_mapping_scheme() -> None:
    session = build_project_session(
        mapping=_mapping(),
        movement_code_scheme=MOVEMENT_SCHEME_V2,
        detected_sheet_names=["North mainline"],
    )
    loaded = session_from_json(session_to_json_bytes(session))

    assert loaded.session["mapping"]["movement_code_scheme"] == MOVEMENT_SCHEME_V2


def test_mapping_excel_workflow_still_round_trips() -> None:
    excel_bytes = mapping_to_excel_bytes(_mapping())
    loaded_excel_mapping = read_mapping_excel(BytesIO(excel_bytes))
    applied = apply_saved_mapping_to_sheets(["North mainline", "East"], loaded_excel_mapping)

    assert applied["raw_sheet"].tolist() == ["North mainline", "East"]
    assert applied.loc[applied["raw_sheet"] == "North mainline", "movement_code"].iloc[0] == "NS"
    assert not bool(applied.loc[applied["raw_sheet"] == "East", "include_in_peak"].iloc[0])


def test_v2_mapping_excel_metadata_detects_approach_movement_and_validates_codes() -> None:
    loaded = read_mapping_excel_with_metadata(V2_MAPPING_XLSX)

    assert loaded.movement_code_scheme == MOVEMENT_SCHEME_V2
    assert loaded.mapping["movement_code"].tolist() == ["NT", "NT", "WT", "ET"]
    assert loaded.mapping["approach_direction"].tolist() == ["N", "N", "W", "E"]
    assert loaded.mapping["movement_type"].tolist() == ["T", "T", "T", "T"]


def test_v1_mapping_excel_without_metadata_defaults_to_from_to() -> None:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        _mapping().to_excel(writer, sheet_name="Mapping", index=False)
    buffer.seek(0)

    loaded = read_mapping_excel_with_metadata(buffer)

    assert loaded.movement_code_scheme == MOVEMENT_SCHEME_V1
    assert loaded.mapping["movement_code"].tolist() == ["NS", "NS", "EW"]


def test_processing_compatibility_guard_distinguishes_v1_and_v2() -> None:
    assert mapping_is_process_compatible(MOVEMENT_SCHEME_V1)
    assert mapping_processing_block_reason(MOVEMENT_SCHEME_V1) == ""
    assert not mapping_is_process_compatible(MOVEMENT_SCHEME_V2)
    assert mapping_processing_block_reason(MOVEMENT_SCHEME_V2) == APPROACH_MOVEMENT_PROCESSING_BLOCK_REASON


def test_single_file_processing_guard_blocks_v2_before_pipeline_work() -> None:
    with pytest.raises(ValueError, match="approach_movement v2"):
        process_tmc(
            raw_sheets={},
            mapping=pd.DataFrame(),
            setup={"movement_code_scheme": MOVEMENT_SCHEME_V2},
            detected_sheets=[],
            generate_workbook=False,
        )


def test_apply_saved_mapping_preserves_optional_review_remark() -> None:
    applied = apply_saved_mapping_to_sheets(["North mainline"], _mapping())

    assert applied.loc[applied["raw_sheet"] == "North mainline", "remark"].iloc[0] == "primary stream"
