from __future__ import annotations

import json
from io import BytesIO

import pandas as pd

import app
from tmc_processor.excel_com_export import ExcelComStatus
from tmc_processor.exporter import template_paths_for_movement_scheme
from tmc_processor.mapping import mapping_to_excel_bytes, normalize_approach_movement_mapping, validate_mapping_scheme
from tmc_processor.mapping_preset import load_mapping_preset
from tmc_processor.mapping_preset import MAPPING_PRESET_TYPE
from tmc_processor.movement_scheme import APPROACH_MOVEMENT_CODES, MOVEMENT_SCHEME_V1, MOVEMENT_SCHEME_V2
from tmc_processor.report_template import (
    DEFAULT_TEMPLATE_MAP_PATH,
    DEFAULT_TEMPLATE_PATH,
    V2_TEMPLATE_MAP_PATH,
    V2_TEMPLATE_PATH,
)
from tmc_processor.session import build_project_session, session_from_json, session_to_json


BASIC_MAPPING_COLUMNS = [
    "raw_sheet",
    "raw_direction",
    "source_stream",
    "movement_code",
    "raw_movement_label",
    "include_in_report",
    "include_in_peak",
]


def _mapping(codes: list[str]) -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "raw_sheet": f"Sheet {index}",
                "raw_direction": "N",
                "movement_code": code,
                "source_stream": "mainline",
                "raw_movement_label": code,
                "from_leg": "N",
                "to_leg": "S",
                "turn_type": "through",
                "facility_type": "at_grade",
                "include_in_peak": True,
                "include_in_report": True,
                "aggregation_method": "sum",
            }
            for index, code in enumerate(codes, start=1)
        ]
    )


def test_old_mapping_preset_without_scheme_defaults_to_from_to() -> None:
    raw = {
        "preset_type": MAPPING_PRESET_TYPE,
        "preset_name": "Old preset",
        "mapping_rows": [
            {"raw_sheet": "N", "output_movement_code": "NE", "include_in_report": True},
        ],
    }

    loaded = load_mapping_preset(json.dumps(raw).encode("utf-8"))

    assert loaded.preset["movement_code_scheme"] == MOVEMENT_SCHEME_V1


def test_mapping_editor_basic_columns_are_scheme_aware() -> None:
    frame = _mapping(["NT"])

    v1 = app._mapping_editor_frame(frame, "Basic", MOVEMENT_SCHEME_V1)
    v2 = app._mapping_editor_frame(frame, "Basic", MOVEMENT_SCHEME_V2)

    assert v1.columns.tolist() == BASIC_MAPPING_COLUMNS
    assert v2.columns.tolist() == BASIC_MAPPING_COLUMNS
    assert "approach_direction" not in v1.columns
    assert "approach_direction" not in v2.columns
    assert "movement_type" not in v2.columns


def test_from_to_basic_columns_remain_unchanged() -> None:
    basic = app._mapping_editor_frame(_mapping(["NE"]), "Basic", MOVEMENT_SCHEME_V1)

    assert basic.columns.tolist() == BASIC_MAPPING_COLUMNS


def test_mapping_editor_basic_labels_match_across_schemes() -> None:
    v1_labels = app._mapping_editor_labels(MOVEMENT_SCHEME_V1)
    v2_labels = app._mapping_editor_labels(MOVEMENT_SCHEME_V2)

    for labels in (v1_labels, v2_labels):
        assert labels["raw_sheet"] == "Sheet ต้นทาง"
        assert labels["raw_direction"] == "ทิศ/stream ต้นทาง"
        assert labels["source_stream"] == "source_stream"
        assert labels["movement_code"] == "รหัส movement"
        assert labels["raw_movement_label"] == "ป้ายแสดงผล"
        assert labels["include_in_report"] == "แสดงในรายงาน"
        assert labels["include_in_peak"] == "ใช้คำนวณ Peak"


def test_default_preview_mapping_rows_do_not_lock_scheme_selector() -> None:
    state = {
        "mapping_table": _mapping([""]).to_dict("records"),
        "tmc_mapping_source": app.MAPPING_SOURCE_DEFAULT_PREVIEW,
    }

    assert app._mapping_rows_are_committed(state) is False


def test_selected_approach_movement_without_committed_rows_keeps_selector_enabled() -> None:
    state = {"tmc_mapping_code_scheme": MOVEMENT_SCHEME_V2}

    assert app._mapping_rows_are_committed(state) is False


def test_loaded_mapping_sources_lock_scheme_selector() -> None:
    rows = _mapping(["NT"]).to_dict("records")

    assert app._mapping_rows_are_committed({"mapping_table": rows, "tmc_mapping_source": app.MAPPING_SOURCE_MAPPING_PRESET})
    assert app._mapping_rows_are_committed({"mapping_table": rows, "tmc_mapping_source": app.MAPPING_SOURCE_MAPPING_EXCEL})


def test_clear_mapping_for_scheme_change_keeps_detected_sheets_and_upload_identity() -> None:
    state: dict[str, object] = {
        "mapping_table": _mapping(["NT"]).to_dict("records"),
        "tmc_mapping_source": app.MAPPING_SOURCE_MAPPING_PRESET,
        "tmc_mapping_preset_upload_identity": ("preset.json", "abc"),
        "tmc_mapping_preset_apply_info": {"matched": 1},
        "tmc_mapping_preset_warnings": ["warning"],
        "tmc_uploaded_identity": ("raw.xlsx", 123),
        "tmc_detected_sheet_names": ["Sheet 1"],
        "mapping_editor_version": 4,
    }

    app._clear_mapping_for_scheme_change(state)

    assert "mapping_table" not in state
    assert state["tmc_mapping_source"] == app.MAPPING_SOURCE_DEFAULT_PREVIEW
    assert state["mapping_editor_version"] == 5
    assert state["tmc_uploaded_identity"] == ("raw.xlsx", 123)
    assert state["tmc_detected_sheet_names"] == ["Sheet 1"]
    assert state["tmc_mapping_preset_ignored_identity"] == ("preset.json", "abc")
    assert "tmc_mapping_preset_apply_info" not in state


def test_project_session_mapping_source_counts_as_committed() -> None:
    session = build_project_session(mapping=_mapping(["NT"]), movement_code_scheme=MOVEMENT_SCHEME_V2)
    loaded = session_from_json(session_to_json(session)).session
    state: dict[str, object] = {}

    app.apply_session_to_state(loaded, state)

    assert state["tmc_mapping_source"] == app.MAPPING_SOURCE_PROJECT_SESSION
    assert app._mapping_rows_are_committed(state)


def test_mapping_excel_download_metadata_is_scheme_specific() -> None:
    v2_bytes = mapping_to_excel_bytes(_mapping(["WT"]), movement_code_scheme=MOVEMENT_SCHEME_V2)
    v1_bytes = mapping_to_excel_bytes(_mapping(["NE"]), movement_code_scheme=MOVEMENT_SCHEME_V1)

    v2_metadata = pd.read_excel(BytesIO(v2_bytes), sheet_name="Metadata")
    v2_mapping = pd.read_excel(BytesIO(v2_bytes), sheet_name="Mapping")
    v1_metadata = pd.read_excel(BytesIO(v1_bytes), sheet_name="Metadata")

    assert dict(zip(v2_metadata["field"], v2_metadata["value"]))["movement_code_scheme"] == MOVEMENT_SCHEME_V2
    assert {"approach_direction", "movement_type"}.issubset(v2_mapping.columns)
    assert v2_mapping.loc[0, "approach_direction"] == "W"
    assert v2_mapping.loc[0, "movement_type"] == "T"
    assert dict(zip(v1_metadata["field"], v1_metadata["value"]))["movement_code_scheme"] == MOVEMENT_SCHEME_V1


def test_approach_movement_mapping_derives_components_from_movement_code() -> None:
    normalized = normalize_approach_movement_mapping(_mapping(["EU"]))

    assert normalized.loc[0, "approach_direction"] == "E"
    assert normalized.loc[0, "movement_type"] == "U"


def test_approach_movement_mapping_overwrites_conflicting_component_values() -> None:
    frame = _mapping(["NT"])
    frame["approach_direction"] = "S"
    frame["movement_type"] = "R"

    normalized = normalize_approach_movement_mapping(frame)

    assert normalized.loc[0, "approach_direction"] == "N"
    assert normalized.loc[0, "movement_type"] == "T"
    assert validate_mapping_scheme(frame, MOVEMENT_SCHEME_V2) == []


def test_approach_movement_editor_options_are_strict_v2_codes() -> None:
    options = app._movement_code_options_for_scheme(pd.DataFrame({"movement_code": [""]}), MOVEMENT_SCHEME_V2)

    assert options == ["", *APPROACH_MOVEMENT_CODES]
    assert {"NE", "NS", "WE", "EN"}.isdisjoint(options)


def test_from_to_editor_options_are_from_to_compatible() -> None:
    options = app._movement_code_options_for_scheme(pd.DataFrame({"movement_code": [""]}), MOVEMENT_SCHEME_V1)

    assert options == ["", *app.MOVEMENT_CODE_OPTIONS]
    assert {"NS", "WE", "EW"}.issubset(options)
    assert set(APPROACH_MOVEMENT_CODES) != set(options) - {""}


def test_approach_movement_valid_and_invalid_codes_are_scheme_specific() -> None:
    assert validate_mapping_scheme(_mapping(["NL", "NT", "NR", "NU"]), MOVEMENT_SCHEME_V2) == []

    issues = validate_mapping_scheme(_mapping(["NE", "NS", "WE", "EN"]), MOVEMENT_SCHEME_V2)

    assert issues
    assert all("invalid approach_movement" in issue for issue in issues)
    assert validate_mapping_scheme(_mapping(["NE"]), MOVEMENT_SCHEME_V1) == []


def test_single_file_export_default_is_template_when_excel_com_available() -> None:
    status = ExcelComStatus(available=True, reason="ok")

    assert app._default_single_export_mode(status) == app.EXCEL_TEMPLATE_EXPORT_MODE
    assert (
        app._coerce_export_mode(None, app._single_export_mode_options(status), app._default_single_export_mode(status))
        == app.EXCEL_TEMPLATE_EXPORT_MODE
    )


def test_single_file_export_default_is_safe_png_when_excel_com_unavailable() -> None:
    status = ExcelComStatus(available=False, reason="missing")

    assert app._single_export_mode_options(status) == [app.SAFE_PNG_EXPORT_MODE]
    assert app._default_single_export_mode(status) == app.SAFE_PNG_EXPORT_MODE
    assert (
        app._coerce_export_mode(
            app.EXCEL_TEMPLATE_EXPORT_MODE,
            app._single_export_mode_options(status),
            app._default_single_export_mode(status),
        )
        == app.SAFE_PNG_EXPORT_MODE
    )


def test_approach_movement_batch_export_default_is_safe_png() -> None:
    status = ExcelComStatus(available=True, reason="ok")

    assert app._batch_export_mode_options(status, MOVEMENT_SCHEME_V2) == [app.BATCH_SAFE_PNG_EXPORT_LABEL]
    assert app._default_batch_export_mode(status, MOVEMENT_SCHEME_V2) == app.BATCH_SAFE_PNG_EXPORT_LABEL


def test_export_template_paths_are_scheme_specific() -> None:
    assert template_paths_for_movement_scheme(MOVEMENT_SCHEME_V1) == (DEFAULT_TEMPLATE_PATH, DEFAULT_TEMPLATE_MAP_PATH)
    assert template_paths_for_movement_scheme(MOVEMENT_SCHEME_V2) == (V2_TEMPLATE_PATH, V2_TEMPLATE_MAP_PATH)


def test_project_session_defaults_and_preserves_movement_code_scheme() -> None:
    old = session_from_json(json.dumps({"mapping": {"rows": _mapping(["NE"]).to_dict("records")}})).session
    assert old["mapping"]["movement_code_scheme"] == MOVEMENT_SCHEME_V1

    session = build_project_session(mapping=_mapping(["NT"]), movement_code_scheme=MOVEMENT_SCHEME_V2)
    round_trip = session_from_json(session_to_json(session)).session

    assert round_trip["mapping"]["movement_code_scheme"] == MOVEMENT_SCHEME_V2


def test_mapping_readiness_status_includes_scheme_and_invalid_code_state() -> None:
    assert app._mapping_readiness_status(_mapping(["NT"]), MOVEMENT_SCHEME_V2, 0) == "approach_movement - พร้อมใช้งาน"
    assert app._mapping_readiness_status(_mapping(["NE"]), MOVEMENT_SCHEME_V2, 1) == "approach_movement - ตรวจพบรหัสไม่ถูกต้อง"
    assert app._mapping_readiness_status(pd.DataFrame(), MOVEMENT_SCHEME_V1, 0) == "from_to - รอกำหนดทิศทาง"
