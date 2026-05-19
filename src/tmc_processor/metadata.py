"""Central setup metadata helpers for TMC reports."""

from __future__ import annotations

from dataclasses import asdict, dataclass
from typing import Any, Mapping

import pandas as pd


METADATA_FIELDS = (
    "project_name",
    "survey_point",
    "survey_date_text",
    "weather",
    "responsible_party",
    "survey_period",
    "tmc_title",
    "north_label",
    "south_label",
    "east_label",
    "west_label",
    "north_road",
    "south_road",
    "east_road",
    "west_road",
    "caption_text",
)

DEFAULT_RESPONSIBLE_PARTY = "ที่ปรึกษา"
DEFAULT_SURVEY_PERIOD = "7.00 - 19.00"
DEFAULT_WEATHER = "แจ่มใส"
DEFAULT_CAPTION_TEXT = "ปริมาณจราจรบนทางแยก PCU/12 ชม. (07.00-19.00 น.)"


@dataclass(frozen=True)
class SetupMetadata:
    project_name: str = ""
    survey_point: str = ""
    survey_date_text: str = ""
    weather: str = ""
    responsible_party: str = ""
    survey_period: str = ""
    tmc_title: str = ""
    north_label: str = ""
    south_label: str = ""
    east_label: str = ""
    west_label: str = ""
    north_road: str = ""
    south_road: str = ""
    east_road: str = ""
    west_road: str = ""
    caption_text: str = ""

    @classmethod
    def from_mapping(cls, setup: Mapping[str, Any] | None) -> "SetupMetadata":
        setup = setup or {}
        return cls(
            project_name=_first_present(setup, "project_name", "project"),
            survey_point=_first_present(setup, "survey_point"),
            survey_date_text=_first_present(setup, "survey_date_text", "survey_date"),
            weather=_first_present(setup, "weather"),
            responsible_party=_first_present(setup, "responsible_party"),
            survey_period=_first_present(setup, "survey_period", "survey_period_text"),
            tmc_title=_first_present(setup, "tmc_title", "tmc_name", "intersection_name"),
            north_label=_first_present(setup, "north_label"),
            south_label=_first_present(setup, "south_label"),
            east_label=_first_present(setup, "east_label"),
            west_label=_first_present(setup, "west_label"),
            north_road=_first_present(setup, "north_road"),
            south_road=_first_present(setup, "south_road"),
            east_road=_first_present(setup, "east_road"),
            west_road=_first_present(setup, "west_road"),
            caption_text=_first_present(setup, "caption_text"),
        )

    @classmethod
    def with_ui_defaults(cls, filename_default: str = "") -> "SetupMetadata":
        return cls(
            survey_point=filename_default,
            weather=DEFAULT_WEATHER,
            responsible_party=DEFAULT_RESPONSIBLE_PARTY,
            survey_period=DEFAULT_SURVEY_PERIOD,
            tmc_title=filename_default,
            caption_text=DEFAULT_CAPTION_TEXT,
        )

    def to_setup_dict(self, extra: Mapping[str, Any] | None = None) -> dict[str, Any]:
        values = asdict(self)
        result = dict(extra or {})
        result.update(values)
        result.update(
            {
                "project": self.project_name,
                "tmc_name": self.tmc_title,
                "intersection_name": self.tmc_title,
                "survey_date": self.survey_date_text,
                "survey_period_text": self.survey_period,
            }
        )
        return result


def setup_with_metadata(setup: Mapping[str, Any] | None) -> dict[str, Any]:
    setup = dict(setup or {})
    return SetupMetadata.from_mapping(setup).to_setup_dict(extra=setup)


def metadata_cell_values(setup: Mapping[str, Any] | None) -> dict[str, Any]:
    values = setup_with_metadata(setup)
    title = _first_present(values, "report_title", "tmc_title")
    survey_point = _survey_point(values)
    if not title and survey_point:
        title = (
            "ข้อมูลปริมาณจราจรบริเวณทางแยก (Turning Movement Count)\n"
            f"จุดนับรถที่ {survey_point}"
        )
    return {
        "report_title": _blank_if_missing(title),
        "project": _blank_if_missing(values.get("project_name")),
        "survey_point": _blank_if_missing(survey_point),
        "survey_date": _blank_if_missing(values.get("survey_date_text")),
        "weather": _blank_if_missing(values.get("weather")),
        "responsible_party": _blank_if_missing(values.get("responsible_party")),
        "survey_period": _blank_if_missing(values.get("survey_period")),
    }


def direction_label_value(setup: Mapping[str, Any] | None, direction: str) -> str:
    values = setup_with_metadata(setup)
    if direction == "south_or_west":
        south = _first_present(values, "south_label")
        west = _first_present(values, "west_label")
        if south and west and str(south) != str(west):
            return f"{south} / {west}"
        return str(south or west or "")
    return str(_first_present(values, f"{direction}_label") or "")


def _blank_if_missing(value: Any) -> Any:
    if value is None:
        return ""
    if isinstance(value, float) and pd.isna(value):
        return ""
    return value


def _first_present(setup: Mapping[str, Any], *keys: str) -> str:
    for key in keys:
        value = setup.get(key)
        if value is None:
            continue
        if isinstance(value, float) and pd.isna(value):
            continue
        if isinstance(value, str):
            if value.strip() == "":
                continue
            return value.strip()
        return str(value)
    return ""


def _survey_point(setup: Mapping[str, Any]) -> str:
    explicit = _first_present(setup, "survey_point")
    if explicit:
        return explicit
    fallback = " ".join(_first_present(setup, key) for key in ("tmc_id", "tmc_title")).strip()
    return fallback or _first_present(setup, "tmc_title")
