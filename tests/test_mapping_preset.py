from __future__ import annotations

from io import BytesIO

import pandas as pd

from tmc_processor.mapping import apply_saved_mapping_to_sheets, mapping_to_excel_bytes, read_mapping_excel
from tmc_processor.mapping_preset import (
    MAPPING_PRESET_TYPE,
    apply_mapping_preset_to_detected_sheets,
    build_mapping_preset,
    load_mapping_preset,
    serialize_mapping_preset,
)
from tmc_processor.session import apply_session_to_state, build_project_session, session_from_json, session_to_json_bytes


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


def test_mapping_preset_load_round_trips_stream_and_aggregation_fields() -> None:
    loaded = load_mapping_preset(serialize_mapping_preset(build_mapping_preset(_mapping()))).preset
    rows = loaded["mapping_rows"]

    assert rows[1]["source_stream"] == "frontage"
    assert rows[1]["aggregation_method"] == "sum"
    assert rows[1]["raw_movement_label"] == "North frontage through"
    assert rows[2]["include_in_peak"] is False


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


def test_mapping_excel_workflow_still_round_trips() -> None:
    excel_bytes = mapping_to_excel_bytes(_mapping())
    loaded_excel_mapping = read_mapping_excel(BytesIO(excel_bytes))
    applied = apply_saved_mapping_to_sheets(["North mainline", "East"], loaded_excel_mapping)

    assert applied["raw_sheet"].tolist() == ["North mainline", "East"]
    assert applied.loc[applied["raw_sheet"] == "North mainline", "movement_code"].iloc[0] == "NS"
    assert not bool(applied.loc[applied["raw_sheet"] == "East", "include_in_peak"].iloc[0])


def test_apply_saved_mapping_preserves_optional_review_remark() -> None:
    applied = apply_saved_mapping_to_sheets(["North mainline"], _mapping())

    assert applied.loc[applied["raw_sheet"] == "North mainline", "remark"].iloc[0] == "primary stream"
