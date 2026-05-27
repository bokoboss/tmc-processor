from __future__ import annotations

import json

import pandas as pd

import app
from tmc_processor.exporter import template_paths_for_movement_scheme
from tmc_processor.mapping import validate_mapping_scheme
from tmc_processor.mapping_preset import load_mapping_preset
from tmc_processor.mapping_preset import MAPPING_PRESET_TYPE
from tmc_processor.movement_scheme import MOVEMENT_SCHEME_V1, MOVEMENT_SCHEME_V2
from tmc_processor.report_template import (
    DEFAULT_TEMPLATE_MAP_PATH,
    DEFAULT_TEMPLATE_PATH,
    V2_TEMPLATE_MAP_PATH,
    V2_TEMPLATE_PATH,
)
from tmc_processor.session import build_project_session, session_from_json, session_to_json


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

    assert "raw_movement_label" in v1.columns
    assert "approach_direction" not in v1.columns
    assert {"approach_direction", "movement_type", "movement_code"}.issubset(v2.columns)
    assert app._mapping_editor_labels(MOVEMENT_SCHEME_V2)["movement_code"] in {"รหัส movement", "movement_code"}


def test_approach_movement_valid_and_invalid_codes_are_scheme_specific() -> None:
    assert validate_mapping_scheme(_mapping(["NL", "NT", "NR", "NU"]), MOVEMENT_SCHEME_V2) == []

    issues = validate_mapping_scheme(_mapping(["NE", "NS", "WE", "EN"]), MOVEMENT_SCHEME_V2)

    assert issues
    assert all("invalid approach_movement" in issue for issue in issues)
    assert validate_mapping_scheme(_mapping(["NE"]), MOVEMENT_SCHEME_V1) == []


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
