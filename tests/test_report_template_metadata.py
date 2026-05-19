from __future__ import annotations

from tmc_processor.metadata import direction_label_value, metadata_cell_values
from tmc_processor.report_template import _direction_label_value, _metadata_values


def test_report_template_metadata_values_match_central_metadata_helper() -> None:
    setup = {
        "project_name": "Demo Project",
        "survey_point": "Demo Point",
        "survey_date_text": "2026-01-01",
        "weather": "Clear",
        "responsible_party": "Team",
        "survey_period": "07:00-19:00",
        "tmc_title": "Demo Intersection",
    }

    assert _metadata_values(setup) == metadata_cell_values(setup)


def test_report_template_direction_label_matches_central_metadata_helper() -> None:
    setup = {
        "north_label": "North",
        "south_label": "South",
        "west_label": "West",
    }

    assert _direction_label_value(setup, "north") == direction_label_value(setup, "north")
    assert _direction_label_value(setup, "south_or_west") == direction_label_value(setup, "south_or_west")
