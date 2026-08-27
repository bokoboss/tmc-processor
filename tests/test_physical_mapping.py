from __future__ import annotations

from io import BytesIO
from types import SimpleNamespace

import pandas as pd
import streamlit as st

import app
from tmc_processor.mapping import (
    PHYSICAL_MOVEMENT_OPTIONS,
    apply_saved_mapping_to_sheets,
    canonical_movement_code_from_physical,
    default_mapping_for_sheets,
    mapping_to_excel_bytes,
    normalize_physical_movement,
    read_mapping_excel,
    physical_mapping_from_canonical_code,
)
from tmc_processor.mapping_preset import (
    apply_mapping_preset_to_detected_sheets,
    build_mapping_preset,
    load_mapping_preset,
    serialize_mapping_preset,
)
from tmc_processor.movement_scheme import MOVEMENT_SCHEME_V1, MOVEMENT_SCHEME_V2


EXPECTED_FROM_TO_CODES = {
    "N": {"Left": "NE", "Through": "NS", "Right": "NW", "U-turn": "NU"},
    "E": {"Left": "ES", "Through": "EW", "Right": "EN", "U-turn": "EU"},
    "S": {"Left": "SW", "Through": "SN", "Right": "SE", "U-turn": "SU"},
    "W": {"Left": "WN", "Through": "WE", "Right": "WS", "U-turn": "WU"},
}


def test_physical_semantics_derive_all_from_to_movements() -> None:
    for approach, movements in EXPECTED_FROM_TO_CODES.items():
        for movement, code in movements.items():
            assert canonical_movement_code_from_physical(approach, movement) == code


def test_physical_semantics_derive_all_approach_movement_movements() -> None:
    expected_tokens = {"Left": "L", "Through": "T", "Right": "R", "U-turn": "U"}

    for approach in EXPECTED_FROM_TO_CODES:
        for movement, token in expected_tokens.items():
            assert canonical_movement_code_from_physical(approach, movement, MOVEMENT_SCHEME_V2) == f"{approach}{token}"


def test_valid_from_to_codes_hydrate_losslessly() -> None:
    for approach, movements in EXPECTED_FROM_TO_CODES.items():
        for movement, code in movements.items():
            hydrated = physical_mapping_from_canonical_code(code)
            assert hydrated is not None
            assert hydrated.approach == approach
            assert hydrated.movement == movement


def test_valid_approach_movement_codes_hydrate_losslessly() -> None:
    for approach in EXPECTED_FROM_TO_CODES:
        for movement, token in {"Left": "L", "Through": "T", "Right": "R", "U-turn": "U"}.items():
            hydrated = physical_mapping_from_canonical_code(f"{approach}{token}", MOVEMENT_SCHEME_V2)
            assert hydrated is not None
            assert hydrated.approach == approach
            assert hydrated.movement == movement


def test_unknown_or_wrong_scheme_codes_require_advanced_mapping() -> None:
    assert physical_mapping_from_canonical_code("XX") is None
    assert physical_mapping_from_canonical_code("NS", MOVEMENT_SCHEME_V2) is None


def test_physical_movement_accepts_operator_and_legacy_tokens() -> None:
    assert normalize_physical_movement("Left") == "L"
    assert normalize_physical_movement("u_turn") == "U"
    assert normalize_physical_movement("T") == "T"
    assert PHYSICAL_MOVEMENT_OPTIONS == ["", "Left", "Through", "Right", "U-turn"]


def test_physical_builder_rejects_incomplete_selection() -> None:
    for approach, movement in [("", "Through"), ("N", "")]:
        try:
            canonical_movement_code_from_physical(approach, movement)
        except ValueError:
            pass
        else:  # pragma: no cover - assertion guard
            raise AssertionError("incomplete physical selection should be rejected")


def test_physical_helpers_do_not_infer_from_sheet_names() -> None:
    mapping = default_mapping_for_sheets(["1", "North", "NB", "ทิศ 2"])

    assert mapping["movement_code"].tolist() == ["", "", "", ""]
    assert mapping["raw_sheet"].tolist() == ["1", "North", "NB", "ทิศ 2"]


def _mapping_row(code: str = "NS") -> dict[str, object]:
    return {
        "raw_sheet": "source-1",
        "raw_direction": "N",
        "movement_code": code,
        "source_stream": "frontage",
        "raw_movement_label": "survey label",
        "from_leg": "N",
        "to_leg": "S",
        "turn_type": "T",
        "facility_type": "frontage",
        "include_in_peak": True,
        "include_in_report": True,
        "aggregation_method": "sum",
        "note": "retain note",
        "remark": "retain remark",
        "legacy_metadata": "retain metadata",
    }


def test_basic_frame_hydrates_physical_controls_and_hides_advanced_fields() -> None:
    frame = app._mapping_editor_frame(pd.DataFrame([_mapping_row("NE")]), "Basic", MOVEMENT_SCHEME_V1)

    assert frame.loc[0, "physical_approach"] == "N"
    assert frame.loc[0, "physical_movement"] == "Left"
    assert frame.loc[0, "derived_code"] == "NE"
    assert frame.loc[0, "status"] == "Ready · NE"
    assert {"source_stream", "movement_code", "facility_type", "from_leg", "to_leg"}.isdisjoint(frame.columns)


def test_basic_physical_edit_derives_canonical_code_and_existing_normalization_fields() -> None:
    base = default_mapping_for_sheets(["source-1"])
    basic = app._mapping_editor_frame(base, "Basic", MOVEMENT_SCHEME_V1)
    basic.loc[0, "physical_approach"] = "N"
    basic.loc[0, "physical_movement"] = "Right"

    merged = app._merge_mapping_editor_result(base, basic, MOVEMENT_SCHEME_V1)

    assert merged.loc[0, "movement_code"] == "NW"
    assert merged.loc[0, ["from_leg", "to_leg", "turn_type"]].to_dict() == {
        "from_leg": "N",
        "to_leg": "W",
        "turn_type": "R",
    }


def test_basic_physical_edit_preserves_advanced_and_legacy_metadata() -> None:
    base = pd.DataFrame([_mapping_row("NS")])
    basic = app._mapping_editor_frame(base, "Basic", MOVEMENT_SCHEME_V1)
    basic.loc[0, "physical_approach"] = "N"
    basic.loc[0, "physical_movement"] = "Left"

    merged = app._merge_mapping_editor_result(base, basic, MOVEMENT_SCHEME_V1)

    assert merged.loc[0, "movement_code"] == "NE"
    assert merged.loc[0, "source_stream"] == "frontage"
    assert merged.loc[0, "aggregation_method"] == "sum"
    assert merged.loc[0, "facility_type"] == "frontage"
    assert merged.loc[0, "note"] == "retain note"
    assert merged.loc[0, "remark"] == "retain remark"
    assert merged.loc[0, "legacy_metadata"] == "retain metadata"


def test_advanced_only_code_is_preserved_and_marked_advanced_required() -> None:
    base = pd.DataFrame([_mapping_row("custom-code")])

    basic = app._mapping_editor_frame(base, "Basic", MOVEMENT_SCHEME_V1)
    merged = app._merge_mapping_editor_result(base, basic, MOVEMENT_SCHEME_V1)

    assert basic.loc[0, "physical_approach"] == ""
    assert basic.loc[0, "physical_movement"] == ""
    assert basic.loc[0, "derived_code"] == "CUSTOM-CODE"
    assert basic.loc[0, "status"].startswith("Advanced required · source-1")
    assert merged.loc[0, "movement_code"] == "custom-code"
    assert merged.loc[0, "legacy_metadata"] == "retain metadata"


def test_excluded_advanced_only_code_is_still_marked_advanced_required() -> None:
    row = _mapping_row("custom-code")
    row["include_in_report"] = False
    basic = app._mapping_editor_frame(pd.DataFrame([row]), "Basic", MOVEMENT_SCHEME_V1)

    assert basic.loc[0, "status"].startswith("Advanced required")
    assert basic.loc[0, "derived_code"] == "CUSTOM-CODE"


def test_basic_widget_state_accepts_physical_edits() -> None:
    base = default_mapping_for_sheets(["source-1"])
    version = 7
    state = {
        app._basic_mapping_widget_key("physical_approach", version, 0): "E",
        app._basic_mapping_widget_key("physical_movement", version, 0): "Through",
    }

    mapping = app._mapping_from_editor_widget_state(
        base,
        "Basic",
        MOVEMENT_SCHEME_V1,
        app._mapping_editor_key("Basic", version),
        state,
    )

    assert mapping.loc[0, "movement_code"] == "EW"
    assert mapping.loc[0, ["from_leg", "to_leg", "turn_type"]].to_dict() == {
        "from_leg": "E",
        "to_leg": "W",
        "turn_type": "T",
    }


def test_mapping_excel_hydrates_basic_without_changing_canonical_mapping() -> None:
    original = pd.DataFrame([_mapping_row("NE")])
    loaded = read_mapping_excel(BytesIO(mapping_to_excel_bytes(original)))
    applied = apply_saved_mapping_to_sheets(["source-1"], loaded)

    basic = app._mapping_editor_frame(applied, "Basic", MOVEMENT_SCHEME_V1)
    round_tripped = app._merge_mapping_editor_result(applied, basic, MOVEMENT_SCHEME_V1)

    assert basic.loc[0, ["physical_approach", "physical_movement"]].to_dict() == {
        "physical_approach": "N",
        "physical_movement": "Left",
    }
    assert round_tripped.loc[0, "movement_code"] == "NE"
    assert round_tripped.loc[0, "source_stream"] == "frontage"
    assert round_tripped.loc[0, "facility_type"] == "frontage"


def test_mapping_preset_hydrates_basic_without_changing_canonical_mapping() -> None:
    original = pd.DataFrame([_mapping_row("NS")])
    preset = load_mapping_preset(
        serialize_mapping_preset(build_mapping_preset(original, movement_code_scheme=MOVEMENT_SCHEME_V1))
    ).preset
    applied = apply_mapping_preset_to_detected_sheets(preset, ["source-1"]).mapping

    basic = app._mapping_editor_frame(applied, "Basic", MOVEMENT_SCHEME_V1)
    round_tripped = app._merge_mapping_editor_result(applied, basic, MOVEMENT_SCHEME_V1)

    assert basic.loc[0, ["physical_approach", "physical_movement"]].to_dict() == {
        "physical_approach": "N",
        "physical_movement": "Through",
    }
    assert round_tripped.loc[0, "movement_code"] == "NS"
    assert round_tripped.loc[0, "source_stream"] == "frontage"


def test_physical_mapping_edit_invalidates_single_analysis_review_and_export() -> None:
    st.session_state.clear()
    app.initialize_setup_state_once("demo.xlsx")
    st.session_state["mapping_table"] = default_mapping_for_sheets(["source-1"]).to_dict("records")
    st.session_state["tmc_processed"] = {"result": SimpleNamespace()}
    st.session_state["tmc_output"] = {"workbook_bytes": b"old"}
    st.session_state["tmc_selected_pce_factors"] = app._current_pce_factors_from_state()
    app._sync_single_workflow_from_state(
        source_bytes=b"source-a",
        source_file_name="demo.xlsx",
        export_mode=app.SAFE_PNG_EXPORT_MODE,
    )

    base = pd.DataFrame(st.session_state["mapping_table"])
    basic = app._mapping_editor_frame(base, "Basic", MOVEMENT_SCHEME_V1)
    basic.loc[0, "physical_approach"] = "N"
    basic.loc[0, "physical_movement"] = "Through"
    st.session_state["mapping_table"] = app._merge_mapping_editor_result(base, basic).to_dict("records")

    transition = app._sync_single_workflow_from_state(
        source_bytes=b"source-a",
        source_file_name="demo.xlsx",
        export_mode=app.SAFE_PNG_EXPORT_MODE,
    )

    assert transition.analysis_invalidated is True
    assert transition.review_invalidated is True
    assert transition.export_invalidated is True
    assert st.session_state["tmc_pce_results_stale"] is True
    assert "tmc_output" not in st.session_state


def test_basic_advanced_view_toggle_is_view_only_for_workflow_state() -> None:
    st.session_state.clear()
    app.initialize_setup_state_once("demo.xlsx")
    st.session_state["mapping_table"] = default_mapping_for_sheets(["source-1"]).to_dict("records")
    st.session_state["tmc_processed"] = {"result": SimpleNamespace()}
    st.session_state["tmc_output"] = {"workbook_bytes": b"old"}
    st.session_state["tmc_selected_pce_factors"] = app._current_pce_factors_from_state()
    app._sync_single_workflow_from_state(
        source_bytes=b"source-a",
        source_file_name="demo.xlsx",
        export_mode=app.SAFE_PNG_EXPORT_MODE,
    )
    previous = app._workflow_state_for_mode(app.WORKFLOW_SINGLE_MODE)
    assert previous is not None

    st.session_state["mapping_editor_view_mode"] = "Advanced"
    st.session_state["mapping_editor_version"] = 12
    transition = app._sync_single_workflow_from_state(
        source_bytes=b"source-a",
        source_file_name="demo.xlsx",
        export_mode=app.SAFE_PNG_EXPORT_MODE,
    )

    assert transition.changed_fields == ()
    current = app._workflow_state_for_mode(app.WORKFLOW_SINGLE_MODE)
    assert current is not None
    assert current.revisions == previous.revisions
    assert "tmc_processed" in st.session_state
    assert "tmc_output" in st.session_state
