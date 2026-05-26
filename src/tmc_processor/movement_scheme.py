"""Movement scheme constants and helpers."""

from __future__ import annotations

from dataclasses import dataclass


MOVEMENT_SCHEME_V1 = "from_to"
MOVEMENT_SCHEME_V2 = "approach_movement"
MOVEMENT_SCHEMES = [MOVEMENT_SCHEME_V1, MOVEMENT_SCHEME_V2]

APPROACH_DIRECTIONS = ["N", "S", "E", "W"]
MOVEMENT_TYPES = ["L", "T", "R", "U"]

APPROACH_MOVEMENT_CODES = [
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

_APPROACH_MOVEMENT_CODE_SET = set(APPROACH_MOVEMENT_CODES)

_APPROACH_DIRECTION_LABELS = {
    "N": "Northbound",
    "S": "Southbound",
    "E": "Eastbound",
    "W": "Westbound",
}

_MOVEMENT_TYPE_LABELS = {
    "L": "Left turn",
    "T": "Through",
    "R": "Right turn",
    "U": "U-turn",
}


@dataclass(frozen=True)
class ApproachMovement:
    """Parsed v2 approach-movement code."""

    approach_direction: str
    movement_type: str


def _normalize_token(value: object, field_name: str) -> str:
    if not isinstance(value, str):
        raise ValueError(f"{field_name} must be a string.")
    return value.strip()


def is_approach_movement_code(code: object) -> bool:
    """Return whether code is a valid v2 approach-movement code."""

    if not isinstance(code, str):
        return False
    return code.strip() in _APPROACH_MOVEMENT_CODE_SET


def normalize_movement_code_scheme(value: object, *, default: str = MOVEMENT_SCHEME_V1) -> str:
    """Return a validated movement-code scheme name."""

    text = str(value or "").strip()
    scheme = text or default
    if scheme not in MOVEMENT_SCHEMES:
        raise ValueError(f"Unsupported movement_code_scheme: {value!r}.")
    return scheme


def validate_approach_movement_code(code: object) -> str:
    """Validate and return a normalized v2 approach-movement code."""

    normalized_code = _normalize_token(code, "approach-movement code")
    if normalized_code not in _APPROACH_MOVEMENT_CODE_SET:
        raise ValueError(f"Invalid approach-movement code: {code!r}.")
    return normalized_code


def parse_approach_movement_code(code: object) -> ApproachMovement:
    """Parse a v2 approach-movement code into its component tokens."""

    normalized_code = validate_approach_movement_code(code)
    return ApproachMovement(
        approach_direction=normalized_code[0],
        movement_type=normalized_code[1],
    )


def build_approach_movement_code(approach_direction: object, movement_type: object) -> str:
    """Build and validate a v2 approach-movement code from component tokens."""

    normalized_direction = _normalize_token(approach_direction, "approach direction")
    normalized_movement_type = _normalize_token(movement_type, "movement type")
    return validate_approach_movement_code(f"{normalized_direction}{normalized_movement_type}")


def approach_direction_label(direction: object) -> str:
    """Return a readable label for a v2 approach direction token."""

    normalized_direction = _normalize_token(direction, "approach direction")
    try:
        return _APPROACH_DIRECTION_LABELS[normalized_direction]
    except KeyError as exc:
        raise ValueError(f"Invalid approach direction: {direction!r}.") from exc


def movement_type_label(movement_type: object) -> str:
    """Return a readable label for a v2 movement type token."""

    normalized_movement_type = _normalize_token(movement_type, "movement type")
    try:
        return _MOVEMENT_TYPE_LABELS[normalized_movement_type]
    except KeyError as exc:
        raise ValueError(f"Invalid movement type: {movement_type!r}.") from exc


def approach_movement_display_label(code: object) -> str:
    """Return a readable display label for a v2 approach-movement code."""

    movement = parse_approach_movement_code(code)
    direction_label = approach_direction_label(movement.approach_direction)
    movement_label = movement_type_label(movement.movement_type)
    return f"{validate_approach_movement_code(code)} - {direction_label} {movement_label}"
