from __future__ import annotations

import pytest

from tmc_processor.movement_scheme import (
    APPROACH_DIRECTIONS,
    APPROACH_MOVEMENT_CODES,
    MOVEMENT_SCHEME_V1,
    MOVEMENT_SCHEME_V2,
    MOVEMENT_TYPES,
    approach_direction_label,
    approach_movement_display_label,
    build_approach_movement_code,
    derive_movement_leg_mapping_from_code,
    is_approach_movement_code,
    movement_type_label,
    parse_approach_movement_code,
    validate_approach_movement_code,
)


def test_movement_scheme_constants_are_stable() -> None:
    assert MOVEMENT_SCHEME_V1 == "from_to"
    assert MOVEMENT_SCHEME_V2 == "approach_movement"
    assert APPROACH_DIRECTIONS == ["N", "S", "E", "W"]
    assert MOVEMENT_TYPES == ["L", "T", "R", "U"]
    assert APPROACH_MOVEMENT_CODES == [
        "NL",
        "NT",
        "NR",
        "NU",
        "SL",
        "ST",
        "SR",
        "SU",
        "EL",
        "ET",
        "ER",
        "EU",
        "WL",
        "WT",
        "WR",
        "WU",
    ]
    assert len(APPROACH_MOVEMENT_CODES) == 16


@pytest.mark.parametrize("code", APPROACH_MOVEMENT_CODES)
def test_all_approach_movement_codes_validate(code: str) -> None:
    assert is_approach_movement_code(code)
    assert validate_approach_movement_code(code) == code


def test_validate_strips_outer_whitespace() -> None:
    assert validate_approach_movement_code(" NL ") == "NL"


def test_parse_approach_movement_code() -> None:
    parsed = parse_approach_movement_code("NL")

    assert parsed.approach_direction == "N"
    assert parsed.movement_type == "L"


def test_build_approach_movement_code() -> None:
    assert build_approach_movement_code("N", "L") == "NL"
    assert build_approach_movement_code(" N ", " L ") == "NL"


def test_derive_movement_leg_mapping_uses_left_hand_traffic_for_all_v2_codes() -> None:
    expected = {
        "NL": ("N", "E"),
        "NT": ("N", "S"),
        "NR": ("N", "W"),
        "NU": ("N", "N"),
        "EL": ("E", "S"),
        "ET": ("E", "W"),
        "ER": ("E", "N"),
        "EU": ("E", "E"),
        "SL": ("S", "W"),
        "ST": ("S", "N"),
        "SR": ("S", "E"),
        "SU": ("S", "S"),
        "WL": ("W", "N"),
        "WT": ("W", "E"),
        "WR": ("W", "S"),
        "WU": ("W", "W"),
    }

    for code, (from_leg, to_leg) in expected.items():
        derived = derive_movement_leg_mapping_from_code(code)
        assert (derived.from_leg, derived.to_leg) == (from_leg, to_leg)


@pytest.mark.parametrize("code", ["N", "NLL", "NS", "WE", "EN", "XX", "", None])
def test_invalid_codes_fail(code: object) -> None:
    assert not is_approach_movement_code(code)
    with pytest.raises(ValueError):
        validate_approach_movement_code(code)
    with pytest.raises(ValueError):
        parse_approach_movement_code(code)


@pytest.mark.parametrize("code", ["nl", "Nt", "en"])
def test_lowercase_or_mixed_case_codes_fail(code: str) -> None:
    assert not is_approach_movement_code(code)
    with pytest.raises(ValueError):
        validate_approach_movement_code(code)


def test_leg_to_leg_codes_are_not_silently_converted() -> None:
    with pytest.raises(ValueError):
        validate_approach_movement_code("NS")
    with pytest.raises(ValueError):
        build_approach_movement_code("N", "S")


def test_labels_are_stable_and_readable() -> None:
    assert approach_direction_label("N") == "Northbound"
    assert approach_direction_label(" S ") == "Southbound"
    assert movement_type_label("L") == "Left turn"
    assert movement_type_label(" U ") == "U-turn"
    assert approach_movement_display_label("NL") == "NL - Northbound Left turn"


def test_invalid_labels_fail() -> None:
    with pytest.raises(ValueError):
        approach_direction_label("NE")
    with pytest.raises(ValueError):
        movement_type_label("S")
