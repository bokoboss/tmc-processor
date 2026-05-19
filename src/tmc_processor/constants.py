"""Project-wide constants from AGENTS.md."""

VEHICLE_CLASSES = [
    "Bicy",
    "MC",
    "PC<7",
    "PC>7",
    "LB",
    "MB",
    "HB",
    "LT",
    "MT",
    "HT",
    "TR",
    "STR",
]

PCE_FACTORS = {
    "Bicy": 0.0,
    "MC": 0.333,
    "PC<7": 1.0,
    "PC>7": 1.0,
    "LB": 1.5,
    "MB": 1.5,
    "HB": 2.1,
    "LT": 1.0,
    "MT": 2.1,
    "HT": 2.5,
    "TR": 2.5,
    "STR": 2.5,
}

MAPPING_COLUMNS = [
    "raw_sheet",
    "raw_direction",
    "movement_code",
    "source_stream",
    "raw_movement_label",
    "from_leg",
    "to_leg",
    "turn_type",
    "facility_type",
    "include_in_peak",
    "include_in_report",
    "aggregation_method",
]

MOVEMENT_CODE_OPTIONS = [
    "NE",
    "NS",
    "NW",
    "NU",
    "SW",
    "SN",
    "SE",
    "SU",
    "WN",
    "WE",
    "WS",
    "WU",
    "ES",
    "EW",
    "EN",
    "EU",
]

LEG_OPTIONS = ["N", "S", "E", "W", "NE", "NW", "SE", "SW", "CUSTOM"]
SOURCE_STREAM_OPTIONS = ["mainline", "frontage", "service_road", "ramp", "other"]
AGGREGATION_METHOD_OPTIONS = ["sum"]
TURN_TYPE_OPTIONS = ["through", "left", "right", "u_turn", "other"]
FACILITY_TYPE_OPTIONS = ["at_grade", "frontage", "overpass", "underpass", "ramp", "other"]

NORMALIZED_COLUMNS = [
    "project",
    "tmc_id",
    "tmc_name",
    "survey_date",
    "raw_sheet",
    "raw_direction",
    "source_direction",
    "source_stream",
    "raw_movement_label",
    "movement_code",
    "output_movement_code",
    "from_leg",
    "to_leg",
    "turn_type",
    "facility_type",
    "include_in_peak",
    "include_in_report",
    "time_start",
    "time_end",
    "vehicle_class",
    "count",
    "pce_factor",
    "pcu",
]

AM_WINDOW = ("07:00", "12:00")
PM_WINDOW = ("15:00", "19:00")
BASE_INTERVAL_MINUTES = 15
PEAK_INTERVALS = 4
PEAK_MODE_FIXED_HOURLY = "fixed_hourly"
PEAK_MODE_ROLLING_60MIN = "rolling_60min"
PEAK_MODE_OPTIONS = [PEAK_MODE_FIXED_HOURLY, PEAK_MODE_ROLLING_60MIN]
DEFAULT_PEAK_MODE = PEAK_MODE_FIXED_HOURLY
