"""Mapping table helpers."""

from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO
from typing import BinaryIO

import pandas as pd

from .constants import (
    AGGREGATION_METHOD_OPTIONS,
    FACILITY_TYPE_OPTIONS,
    MAPPING_COLUMNS,
    MOVEMENT_CODE_OPTIONS,
    SOURCE_STREAM_OPTIONS,
    TURN_TYPE_OPTIONS,
)
from .importer import extract_raw_direction
from .movement_scheme import (
    APPROACH_MOVEMENT_CODES,
    MOVEMENT_SCHEME_V1,
    MOVEMENT_SCHEME_V2,
    build_approach_movement_code,
    derive_movement_leg_mapping_from_code,
    is_approach_movement_code,
    normalize_movement_code_scheme,
    parse_approach_movement_code,
)


APPROACH_MOVEMENT_PROCESSING_BLOCK_REASON = (
    "Mapping นี้ใช้ระบบรหัส approach_movement v2 ซึ่งยังไม่ได้เปิดใช้ใน pipeline หลัก "
    "โปรดใช้ Mapping แบบ from_to สำหรับการประมวลผลในเวอร์ชันนี้"
)
OPTIONAL_MAPPING_METADATA_COLUMNS = ["approach_direction", "movement_type"]


_SOURCE_STREAM_DEFAULT = "mainline"
_AGGREGATION_METHOD_DEFAULT = "sum"
_FACILITY_TYPE_DEFAULT = "at_grade"

PHYSICAL_APPROACH_OPTIONS = ["", "N", "E", "S", "W"]
PHYSICAL_MOVEMENT_OPTIONS = ["", "Left", "Through", "Right", "U-turn"]

_PHYSICAL_MOVEMENT_ALIASES = {
    "l": "L",
    "left": "L",
    "left_turn": "L",
    "t": "T",
    "through": "T",
    "straight": "T",
    "r": "R",
    "right": "R",
    "right_turn": "R",
    "u": "U",
    "u_turn": "U",
    "uturn": "U",
}
_PHYSICAL_MOVEMENT_LABELS = {
    "L": "Left",
    "T": "Through",
    "R": "Right",
    "U": "U-turn",
}

_SOURCE_STREAM_ALIASES = {
    "mainline": "mainline",
    "frontage": "frontage",
    "service_road": "service_road",
    "service": "service_road",
    "ramp": "ramp",
    "other": "other",
}
_AGGREGATION_METHOD_ALIASES = {"sum": "sum"}
_TURN_TYPE_ALIASES = {
    "l": "L",
    "t": "T",
    "r": "R",
    "u": "U",
    "through": "through",
    "left": "left",
    "right": "right",
    "u_turn": "u_turn",
    "uturn": "u_turn",
    "combined": "other",
    "other": "other",
}
_TURN_TYPE_OPTIONS = [*TURN_TYPE_OPTIONS, "L", "T", "R", "U"]
_FACILITY_TYPE_ALIASES = {
    "at_grade": "at_grade",
    "frontage": "frontage",
    "overpass": "overpass",
    "underpass": "underpass",
    "ramp": "ramp",
    "u_turn": "other",
    "uturn": "other",
    "other": "other",
}


@dataclass(frozen=True)
class MappingExcelLoadResult:
    mapping: pd.DataFrame
    metadata: dict[str, str]
    warnings: tuple[str, ...] = ()

    @property
    def movement_code_scheme(self) -> str:
        return self.metadata.get("movement_code_scheme", MOVEMENT_SCHEME_V1)


@dataclass(frozen=True)
class PhysicalMapping:
    """Lossless operator-facing representation of a supported movement."""

    approach: str
    movement: str


def _choice_key(value: object) -> str:
    return str(value or "").strip().casefold().replace("-", "_").replace(" ", "_")


def _physical_text(value: object) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    return str(value).strip()


def normalize_physical_approach(value: object) -> str:
    """Return a supported physical approach token or an empty value."""

    approach = _physical_text(value).upper()
    return approach if approach in PHYSICAL_APPROACH_OPTIONS else ""


def normalize_physical_movement(value: object) -> str:
    """Return the canonical movement token (L/T/R/U) or an empty value."""

    return _PHYSICAL_MOVEMENT_ALIASES.get(_choice_key(_physical_text(value)), "")


def physical_movement_label(value: object) -> str:
    """Return the operator-facing label for a movement token or label."""

    return _PHYSICAL_MOVEMENT_LABELS.get(normalize_physical_movement(value), "")


def _canonical_choice(value: object, aliases: dict[str, str], options: list[str], default: str = "") -> str:
    text = str(value or "").strip()
    if not text:
        return default
    canonical = aliases.get(_choice_key(text))
    if canonical in options:
        return canonical
    return "other" if "other" in options else default


def _canonical_movement_code(value: object) -> str:
    text = str(value or "").strip()
    if not text:
        return ""
    upper_text = text.upper()
    return upper_text if upper_text in MOVEMENT_CODE_OPTIONS else text


def _unknown_values(series: pd.Series, aliases: dict[str, str], options: list[str]) -> list[str]:
    unknown = []
    for value in series.dropna().astype(str):
        text = value.strip()
        if text and aliases.get(_choice_key(text)) not in options:
            unknown.append(text)
    return sorted(set(unknown))


def mapping_control_warnings(mapping: pd.DataFrame, movement_code_scheme: str = MOVEMENT_SCHEME_V1) -> list[str]:
    """Return non-blocking warnings for legacy values outside editor dropdowns."""

    scheme = normalize_movement_code_scheme(movement_code_scheme)
    warnings: list[str] = []
    checks = {
        "source_stream": (_SOURCE_STREAM_ALIASES, SOURCE_STREAM_OPTIONS, "other"),
        "turn_type": (_TURN_TYPE_ALIASES, _TURN_TYPE_OPTIONS, "other"),
        "facility_type": (_FACILITY_TYPE_ALIASES, FACILITY_TYPE_OPTIONS, "other"),
        "aggregation_method": (_AGGREGATION_METHOD_ALIASES, AGGREGATION_METHOD_OPTIONS, _AGGREGATION_METHOD_DEFAULT),
    }
    for column, (aliases, options, fallback) in checks.items():
        if column not in mapping.columns:
            continue
        values = _unknown_values(mapping[column], aliases, options)
        if values:
            warnings.append(
                f"Unknown {column} value(s) {', '.join(values)} were loaded safely as {fallback}."
            )

    movement_column = "movement_code" if "movement_code" in mapping.columns else "output_movement_code"
    if movement_column in mapping.columns:
        movement_options = APPROACH_MOVEMENT_CODES if scheme == MOVEMENT_SCHEME_V2 else MOVEMENT_CODE_OPTIONS
        values = sorted(
            {
                text
                for text in mapping[movement_column].dropna().astype(str).str.strip()
                if text and text.upper() not in movement_options
            }
        )
        if values:
            warnings.append(
                f"Legacy movement code value(s) {', '.join(values)} are preserved for compatibility."
            )
    return warnings


def selectbox_options_with_existing_values(base_options: list[str], values: pd.Series | None = None) -> list[str]:
    """Append existing legacy values so Streamlit selectboxes can render old mappings safely."""

    options = list(dict.fromkeys(base_options))
    if values is None:
        return options
    for value in values.dropna().astype(str).str.strip():
        if value and value not in options:
            options.append(value)
    return options


def default_mapping_for_sheets(raw_sheets: list[str]) -> pd.DataFrame:
    rows = []
    for sheet in raw_sheets:
        raw_direction = extract_raw_direction(sheet) or ""
        rows.append(
            {
                "raw_sheet": sheet,
                "raw_direction": raw_direction,
                "movement_code": "",
                "source_stream": _SOURCE_STREAM_DEFAULT,
                "raw_movement_label": raw_direction,
                "from_leg": "",
                "to_leg": "",
                "turn_type": "",
                "facility_type": _FACILITY_TYPE_DEFAULT,
                "include_in_peak": True,
                "include_in_report": True,
                "aggregation_method": _AGGREGATION_METHOD_DEFAULT,
            }
        )
    return pd.DataFrame(rows, columns=MAPPING_COLUMNS)


def _bool_series(series: pd.Series, default: bool) -> pd.Series:
    def parse(value) -> bool:
        if pd.isna(value):
            return default
        if isinstance(value, bool):
            return value
        text = str(value).strip().casefold()
        if text in {"false", "0", "no", "n", "off"}:
            return False
        if text in {"true", "1", "yes", "y", "on"}:
            return True
        return bool(value)

    return series.map(parse)


def clean_mapping(mapping: pd.DataFrame) -> pd.DataFrame:
    cleaned = mapping.copy()
    if "movement_code" not in cleaned.columns and "output_movement_code" in cleaned.columns:
        cleaned["movement_code"] = cleaned["output_movement_code"]
    elif "movement_code" in cleaned.columns and "output_movement_code" in cleaned.columns:
        movement_code = cleaned["movement_code"].fillna("").astype(str).str.strip()
        cleaned.loc[movement_code == "", "movement_code"] = cleaned.loc[movement_code == "", "output_movement_code"]
    for column in MAPPING_COLUMNS:
        if column not in cleaned.columns:
            if column in {"include_in_peak", "include_in_report"}:
                cleaned[column] = True
            elif column == "source_stream":
                cleaned[column] = _SOURCE_STREAM_DEFAULT
            elif column == "aggregation_method":
                cleaned[column] = _AGGREGATION_METHOD_DEFAULT
            elif column == "facility_type":
                cleaned[column] = _FACILITY_TYPE_DEFAULT
            else:
                cleaned[column] = ""
    optional_values = {
        column: cleaned[column].fillna("").astype(str).tolist()
        for column in OPTIONAL_MAPPING_METADATA_COLUMNS
        if column in cleaned.columns
    }
    cleaned = cleaned[MAPPING_COLUMNS]
    cleaned["include_in_peak"] = _bool_series(cleaned["include_in_peak"], default=True)
    cleaned["include_in_report"] = _bool_series(cleaned["include_in_report"], default=True)
    text_columns = [column for column in MAPPING_COLUMNS if not column.startswith("include_")]
    cleaned[text_columns] = cleaned[text_columns].fillna("").astype(str)
    cleaned["movement_code"] = cleaned["movement_code"].map(_canonical_movement_code)
    cleaned["source_stream"] = cleaned["source_stream"].map(
        lambda value: _canonical_choice(value, _SOURCE_STREAM_ALIASES, SOURCE_STREAM_OPTIONS, _SOURCE_STREAM_DEFAULT)
    )
    cleaned.loc[cleaned["raw_movement_label"].str.strip() == "", "raw_movement_label"] = cleaned["raw_direction"]
    cleaned["turn_type"] = cleaned["turn_type"].map(
        lambda value: _canonical_choice(value, _TURN_TYPE_ALIASES, _TURN_TYPE_OPTIONS)
    )
    cleaned["facility_type"] = cleaned["facility_type"].map(
        lambda value: _canonical_choice(value, _FACILITY_TYPE_ALIASES, FACILITY_TYPE_OPTIONS, _FACILITY_TYPE_DEFAULT)
    )
    cleaned["aggregation_method"] = cleaned["aggregation_method"].map(
        lambda value: _canonical_choice(
            value,
            _AGGREGATION_METHOD_ALIASES,
            AGGREGATION_METHOD_OPTIONS,
            _AGGREGATION_METHOD_DEFAULT,
        )
    )
    for column, values in optional_values.items():
        cleaned[column] = values
    return cleaned


def normalize_approach_movement_mapping(mapping: pd.DataFrame) -> pd.DataFrame:
    """Return mapping with v2 component and processor-leg fields derived from movement_code."""

    normalized = clean_mapping(mapping)
    if "approach_direction" not in normalized.columns:
        normalized["approach_direction"] = ""
    if "movement_type" not in normalized.columns:
        normalized["movement_type"] = ""
    for index, code in normalized["movement_code"].fillna("").astype(str).str.strip().items():
        if not is_approach_movement_code(code):
            continue
        parsed = parse_approach_movement_code(code)
        leg_mapping = derive_movement_leg_mapping_from_code(code)
        normalized.at[index, "approach_direction"] = parsed.approach_direction
        normalized.at[index, "movement_type"] = parsed.movement_type
        normalized.at[index, "from_leg"] = leg_mapping.from_leg
        normalized.at[index, "to_leg"] = leg_mapping.to_leg
    return normalized


_FROM_TO_TURN_TYPE = {
    ("N", "E"): "L",
    ("N", "S"): "T",
    ("N", "W"): "R",
    ("N", "N"): "U",
    ("E", "S"): "L",
    ("E", "W"): "T",
    ("E", "N"): "R",
    ("E", "E"): "U",
    ("S", "W"): "L",
    ("S", "N"): "T",
    ("S", "E"): "R",
    ("S", "S"): "U",
    ("W", "N"): "L",
    ("W", "E"): "T",
    ("W", "S"): "R",
    ("W", "W"): "U",
}

_FROM_TO_PHYSICAL_CODE_BY_APPROACH_MOVEMENT = {
    ("N", "L"): "NE",
    ("N", "T"): "NS",
    ("N", "R"): "NW",
    ("N", "U"): "NU",
    ("E", "L"): "ES",
    ("E", "T"): "EW",
    ("E", "R"): "EN",
    ("E", "U"): "EU",
    ("S", "L"): "SW",
    ("S", "T"): "SN",
    ("S", "R"): "SE",
    ("S", "U"): "SU",
    ("W", "L"): "WN",
    ("W", "T"): "WE",
    ("W", "R"): "WS",
    ("W", "U"): "WU",
}
_PHYSICAL_MOVEMENT_BY_FROM_TO_CODE = {
    code: movement for (approach, movement), code in _FROM_TO_PHYSICAL_CODE_BY_APPROACH_MOVEMENT.items()
}


def canonical_movement_code_from_physical(
    approach: object,
    movement: object,
    movement_code_scheme: str = MOVEMENT_SCHEME_V1,
) -> str:
    """Derive a canonical movement code from operator physical semantics.

    ``from_to`` uses the project's existing left-hand-traffic leg semantics;
    ``approach_movement`` keeps the selected approach and movement tokens in
    its existing v2 representation.
    """

    scheme = normalize_movement_code_scheme(movement_code_scheme)
    normalized_approach = normalize_physical_approach(approach)
    normalized_movement = normalize_physical_movement(movement)
    if not normalized_approach:
        raise ValueError("Physical approach is required.")
    if not normalized_movement:
        raise ValueError("Physical movement is required.")
    if scheme == MOVEMENT_SCHEME_V2:
        return build_approach_movement_code(normalized_approach, normalized_movement)
    return _FROM_TO_PHYSICAL_CODE_BY_APPROACH_MOVEMENT[(normalized_approach, normalized_movement)]


def physical_mapping_from_canonical_code(
    movement_code: object,
    movement_code_scheme: str = MOVEMENT_SCHEME_V1,
) -> PhysicalMapping | None:
    """Hydrate Basic physical controls when a canonical code is lossless."""

    scheme = normalize_movement_code_scheme(movement_code_scheme)
    code = _physical_text(movement_code).upper()
    if not code:
        return None
    if scheme == MOVEMENT_SCHEME_V2:
        if not is_approach_movement_code(code):
            return None
        parsed = parse_approach_movement_code(code)
        return PhysicalMapping(parsed.approach_direction, physical_movement_label(parsed.movement_type))
    movement_token = _PHYSICAL_MOVEMENT_BY_FROM_TO_CODE.get(code)
    if movement_token is None:
        return None
    return PhysicalMapping(code[0], physical_movement_label(movement_token))


def canonical_code_from_physical(
    approach: object,
    movement: object,
    movement_code_scheme: str = MOVEMENT_SCHEME_V1,
) -> str:
    """Short alias for :func:`canonical_movement_code_from_physical`."""

    return canonical_movement_code_from_physical(approach, movement, movement_code_scheme)


def physical_mapping_from_code(
    movement_code: object,
    movement_code_scheme: str = MOVEMENT_SCHEME_V1,
) -> PhysicalMapping | None:
    """Short alias for :func:`physical_mapping_from_canonical_code`."""

    return physical_mapping_from_canonical_code(movement_code, movement_code_scheme)


def normalize_from_to_movement_mapping(mapping: pd.DataFrame) -> pd.DataFrame:
    """Return mapping with processor fields derived from NE/NS/NW/NU-style movement_code."""

    normalized = clean_mapping(mapping)
    for index, code in normalized["movement_code"].fillna("").astype(str).str.strip().items():
        if code not in MOVEMENT_CODE_OPTIONS:
            continue
        from_leg = code[0]
        to_leg = from_leg if code[1] == "U" else code[1]
        normalized.at[index, "from_leg"] = from_leg
        normalized.at[index, "to_leg"] = to_leg
        normalized.at[index, "turn_type"] = _FROM_TO_TURN_TYPE[(from_leg, to_leg)]
    return normalized


def normalize_mapping_for_scheme(mapping: pd.DataFrame, movement_code_scheme: str = MOVEMENT_SCHEME_V1) -> pd.DataFrame:
    scheme = normalize_movement_code_scheme(movement_code_scheme)
    return normalize_approach_movement_mapping(mapping) if scheme == MOVEMENT_SCHEME_V2 else normalize_from_to_movement_mapping(mapping)


REQUIRED_MAPPING_FIELDS = ["movement_code", "from_leg", "to_leg", "turn_type", "facility_type"]


def apply_saved_mapping_to_sheets(raw_sheets: list[str], saved_mapping: pd.DataFrame) -> pd.DataFrame:
    """Align a saved mapping workbook to the currently detected raw sheets."""
    current = default_mapping_for_sheets(raw_sheets)
    extra_columns = [
        column
        for column in (*OPTIONAL_MAPPING_METADATA_COLUMNS, "note", "remark")
        if column in saved_mapping.columns and "raw_sheet" in saved_mapping.columns
    ]
    saved_extras = saved_mapping.copy() if extra_columns else pd.DataFrame()
    if extra_columns:
        saved_extras = saved_extras.drop_duplicates("raw_sheet", keep="last").set_index("raw_sheet")
    saved = clean_mapping(saved_mapping)
    saved_by_sheet = saved.drop_duplicates("raw_sheet", keep="last").set_index("raw_sheet")

    rows = []
    for _, current_row in current.iterrows():
        raw_sheet = current_row["raw_sheet"]
        if raw_sheet in saved_by_sheet.index:
            merged = current_row.to_dict()
            saved_row = saved_by_sheet.loc[raw_sheet]
            for column in MAPPING_COLUMNS:
                if column == "raw_sheet":
                    continue
                merged[column] = saved_row[column]
            if extra_columns and raw_sheet in saved_extras.index:
                extra_row = saved_extras.loc[raw_sheet]
                for column in extra_columns:
                    merged[column] = "" if pd.isna(extra_row[column]) else str(extra_row[column])
            merged["raw_direction"] = merged["raw_direction"] or extract_raw_direction(raw_sheet) or ""
            rows.append(merged)
        else:
            rows.append(current_row.to_dict())
    columns = [*MAPPING_COLUMNS, *[column for column in extra_columns if column not in MAPPING_COLUMNS]]
    return pd.DataFrame(rows, columns=columns)


def _read_mapping_excel_metadata(workbook: pd.ExcelFile) -> dict[str, str]:
    if "Metadata" not in workbook.sheet_names:
        return {"movement_code_scheme": MOVEMENT_SCHEME_V1}
    raw = pd.read_excel(workbook, sheet_name="Metadata", header=None)
    metadata: dict[str, str] = {}
    if raw.shape[1] >= 2:
        for _, row in raw.iloc[:, :2].iterrows():
            key = str(row.iloc[0] or "").strip()
            if key and key.casefold() not in {"field", "key"}:
                value = "" if pd.isna(row.iloc[1]) else str(row.iloc[1]).strip()
                metadata[key] = value
    metadata["movement_code_scheme"] = normalize_movement_code_scheme(metadata.get("movement_code_scheme"))
    return metadata


def _mixed_movement_code_issues(cleaned: pd.DataFrame) -> list[str]:
    if "movement_code" not in cleaned:
        return []
    codes = {
        str(code).strip()
        for code in cleaned.loc[cleaned["include_in_report"], "movement_code"].dropna()
        if str(code).strip()
    }
    has_v1 = any(code in MOVEMENT_CODE_OPTIONS and not is_approach_movement_code(code) for code in codes)
    has_v2 = any(is_approach_movement_code(code) and code not in MOVEMENT_CODE_OPTIONS for code in codes)
    if has_v1 and has_v2:
        return ["Mapping contains mixed from_to and approach_movement movement codes."]
    return []


def validate_mapping_scheme(mapping: pd.DataFrame, movement_code_scheme: str = MOVEMENT_SCHEME_V1) -> list[str]:
    scheme = normalize_movement_code_scheme(movement_code_scheme)
    cleaned = normalize_approach_movement_mapping(mapping) if scheme == MOVEMENT_SCHEME_V2 else clean_mapping(mapping)
    issues: list[str] = []
    issues.extend(_mixed_movement_code_issues(cleaned))
    if scheme == MOVEMENT_SCHEME_V1:
        return issues
    included = cleaned[cleaned["include_in_report"]] if "include_in_report" in cleaned else cleaned
    for row_number, row in enumerate(included.to_dict("records"), start=1):
        code = str(row.get("movement_code") or "").strip()
        if not code:
            continue
        if not is_approach_movement_code(code):
            issues.append(f"Row {row_number} has invalid approach_movement output_movement_code {code!r}.")
            continue
    return issues


def mapping_processing_block_reason(movement_code_scheme: str = MOVEMENT_SCHEME_V1) -> str:
    scheme = normalize_movement_code_scheme(movement_code_scheme)
    return APPROACH_MOVEMENT_PROCESSING_BLOCK_REASON if scheme == MOVEMENT_SCHEME_V2 else ""


def mapping_is_process_compatible(movement_code_scheme: str = MOVEMENT_SCHEME_V1) -> bool:
    return not mapping_processing_block_reason(movement_code_scheme)


def read_mapping_excel_with_metadata(excel_file: str | BinaryIO | BytesIO) -> MappingExcelLoadResult:
    workbook = pd.ExcelFile(excel_file)
    sheet_name = "Mapping" if "Mapping" in workbook.sheet_names else workbook.sheet_names[0]
    metadata = _read_mapping_excel_metadata(workbook)
    raw_mapping = pd.read_excel(workbook, sheet_name=sheet_name)
    mapping = (
        normalize_approach_movement_mapping(raw_mapping)
        if metadata["movement_code_scheme"] == MOVEMENT_SCHEME_V2
        else clean_mapping(raw_mapping)
    )
    issues = validate_mapping_scheme(mapping, metadata["movement_code_scheme"])
    if issues:
        raise ValueError("; ".join(issues))
    return MappingExcelLoadResult(mapping=mapping, metadata=metadata)


def read_mapping_excel(excel_file: str | BinaryIO | BytesIO) -> pd.DataFrame:
    return read_mapping_excel_with_metadata(excel_file).mapping


def mapping_to_excel_bytes(mapping: pd.DataFrame, movement_code_scheme: str = MOVEMENT_SCHEME_V1) -> bytes:
    scheme = normalize_movement_code_scheme(movement_code_scheme)
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        cleaned = normalize_approach_movement_mapping(mapping) if scheme == MOVEMENT_SCHEME_V2 else clean_mapping(mapping)
        if scheme == MOVEMENT_SCHEME_V2:
            for column in OPTIONAL_MAPPING_METADATA_COLUMNS:
                if column not in cleaned.columns:
                    cleaned[column] = ""
        cleaned.to_excel(writer, sheet_name="Mapping", index=False)
        pd.DataFrame(
            [
                {"field": "movement_code_scheme", "value": scheme},
            ]
        ).to_excel(writer, sheet_name="Metadata", index=False)
        worksheet = writer.sheets["Mapping"]
        worksheet.freeze_panes = "A2"
        for column_cells in worksheet.columns:
            max_length = max(len(str(cell.value)) if cell.value is not None else 0 for cell in column_cells)
            worksheet.column_dimensions[column_cells[0].column_letter].width = min(max(max_length + 2, 12), 36)
    return buffer.getvalue()


def validate_mapping_for_processing(detected_sheets: list[str], mapping: pd.DataFrame) -> pd.DataFrame:
    cleaned = normalize_from_to_movement_mapping(mapping)
    issues = []

    for sheet in detected_sheets:
        sheet_rows = cleaned[cleaned["raw_sheet"] == sheet]
        included_rows = sheet_rows[sheet_rows["include_in_report"]] if not sheet_rows.empty else sheet_rows
        if sheet_rows.empty:
            missing = REQUIRED_MAPPING_FIELDS
            for field in missing:
                issues.append(
                    {
                        "raw_sheet": sheet,
                        "field": field,
                        "message": f"Detected raw sheet requires {field} before processing.",
                    }
                )
            continue
        for _, row in included_rows.iterrows():
            missing = [field for field in REQUIRED_MAPPING_FIELDS if str(row[field]).strip() == ""]
            for field in missing:
                issues.append(
                    {
                        "raw_sheet": sheet,
                        "field": field,
                        "message": f"Detected raw sheet requires {field} before processing.",
                    }
                )

    return pd.DataFrame(issues, columns=["raw_sheet", "field", "message"])


def validate_mapping_for_processing_by_scheme(
    detected_sheets: list[str],
    mapping: pd.DataFrame,
    movement_code_scheme: str = MOVEMENT_SCHEME_V1,
) -> pd.DataFrame:
    scheme = normalize_movement_code_scheme(movement_code_scheme)
    if scheme == MOVEMENT_SCHEME_V1:
        return validate_mapping_for_processing(detected_sheets, mapping)

    cleaned = normalize_mapping_for_scheme(mapping, scheme)
    issues = []
    for issue in validate_mapping_scheme(cleaned, scheme):
        issues.append({"raw_sheet": "", "field": "movement_code", "message": issue})

    for sheet in detected_sheets:
        sheet_rows = cleaned[cleaned["raw_sheet"] == sheet]
        included_rows = sheet_rows[sheet_rows["include_in_report"]] if not sheet_rows.empty else sheet_rows
        if sheet_rows.empty or included_rows.empty:
            issues.append(
                {
                    "raw_sheet": sheet,
                    "field": "movement_code",
                    "message": "Detected raw sheet requires an approach_movement mapping before processing.",
                }
            )
            continue
        for _, row in included_rows.iterrows():
            if not str(row.get("movement_code") or "").strip():
                issues.append(
                    {
                        "raw_sheet": sheet,
                        "field": "movement_code",
                        "message": "Detected raw sheet requires movement_code before processing.",
                    }
                )
    return pd.DataFrame(issues, columns=["raw_sheet", "field", "message"])


def movement_aggregation_messages(mapping: pd.DataFrame) -> list[str]:
    """Return non-blocking messages for many-to-one report movement mappings."""

    cleaned = clean_mapping(mapping)
    included = cleaned[cleaned["include_in_report"] & (cleaned["movement_code"].str.strip() != "")]
    counts = included.groupby("movement_code")["raw_sheet"].count()
    return [
        f"{movement_code} is aggregated from {count} source streams."
        for movement_code, count in counts.items()
        if int(count) > 1
    ]
