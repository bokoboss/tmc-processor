"""Summary table calculations."""

from __future__ import annotations

import pandas as pd

from .constants import VEHICLE_CLASSES
from .movement_scheme import APPROACH_MOVEMENT_CODES, MOVEMENT_SCHEME_V1, MOVEMENT_SCHEME_V2, normalize_movement_code_scheme
from .time_utils import time_to_minutes, minutes_to_time


STANDARD_MOVEMENT_ORDER = [
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

THAI_VEHICLE_LABELS = {
    "Bicy": "จักรยาน",
    "MC": "รถจักรยานยนต์และสามล้อเครื่อง",
    "PC<7": "รถยนต์นั่งส่วนบุคคลไม่เกิน 7 คน",
    "PC>7": "รถยนต์นั่งส่วนบุคคลเกิน 7 คน",
    "LB": "รถโดยสารขนาดเล็ก",
    "MB": "รถโดยสารขนาดกลาง",
    "HB": "รถโดยสารขนาดใหญ่",
    "LT": "รถบรรทุกขนาดเล็ก 4 ล้อ",
    "MT": "รถบรรทุกขนาดกลาง 6 ล้อ",
    "HT": "รถบรรทุกขนาดใหญ่ 10 ล้อ",
    "TR": "รถบรรทุกพ่วง",
    "STR": "รถบรรทุกกึ่งพ่วง",
}

DEFAULT_VEHICLE_GROUPS = [
    {"Type": "รถ 2,3 ล้อ", "PCE": "< 1.0", "classes": ["Bicy", "MC"]},
    {"Type": "รถขนาดเล็ก", "PCE": "1.0", "classes": ["PC<7", "PC>7", "LB", "LT"]},
    {"Type": "รถขนาดกลาง", "PCE": "1.1-1.4", "classes": ["MB", "HB", "MT"]},
    {"Type": "รถขนาดใหญ่", "PCE": ">= 1.5", "classes": ["HT", "TR", "STR"]},
]


def _hour_label(hour_minute: int) -> str:
    return f"{minutes_to_time(hour_minute).strftime('%H:%M')}-{minutes_to_time(hour_minute + 60).strftime('%H:%M')}"


def _hourly_frame(normalized: pd.DataFrame) -> pd.DataFrame:
    if normalized.empty:
        return pd.DataFrame(columns=["hour_minute"])
    df = normalized.copy()
    df["minute"] = df["time_start"].map(time_to_minutes)
    df = df.dropna(subset=["minute"])
    df["hour_minute"] = (df["minute"].astype(int) // 60) * 60
    return df


def _movement_columns(mapping: pd.DataFrame, movement_code_scheme: str = MOVEMENT_SCHEME_V1) -> list[str]:
    scheme = normalize_movement_code_scheme(movement_code_scheme)
    standard_order = APPROACH_MOVEMENT_CODES if scheme == MOVEMENT_SCHEME_V2 else STANDARD_MOVEMENT_ORDER
    present = []
    if "movement_code" in mapping:
        present = [str(value).strip() for value in mapping["movement_code"].dropna() if str(value).strip()]
    elif "output_movement_code" in mapping:
        present = [str(value).strip() for value in mapping["output_movement_code"].dropna() if str(value).strip()]
    extras = sorted(code for code in set(present) if code not in standard_order)
    return standard_order + extras


def _with_report_movement_code(normalized: pd.DataFrame) -> pd.DataFrame:
    df = normalized.copy()
    if "output_movement_code" in df.columns:
        output = df["output_movement_code"].fillna("").astype(str).str.strip()
        fallback = df["movement_code"].fillna("").astype(str).str.strip() if "movement_code" in df else ""
        df["movement_code"] = output.where(output != "", fallback)
    return df


def hourly_summary(normalized: pd.DataFrame) -> pd.DataFrame:
    if normalized.empty:
        return pd.DataFrame(columns=["hour_start", "hour_end", "count", "pcu"])
    df = normalized.copy()
    df["minute"] = df["time_start"].map(time_to_minutes)
    df = df.dropna(subset=["minute"])
    df["hour_minute"] = (df["minute"].astype(int) // 60) * 60
    grouped = df.groupby("hour_minute", as_index=False).agg(count=("count", "sum"), pcu=("pcu", "sum"))
    grouped["hour_start"] = grouped["hour_minute"].map(minutes_to_time)
    grouped["hour_end"] = (grouped["hour_minute"] + 60).map(minutes_to_time)
    return grouped[["hour_start", "hour_end", "count", "pcu"]]


def movement_summary(normalized: pd.DataFrame) -> pd.DataFrame:
    columns = [
        "movement_code",
        "from_leg",
        "to_leg",
        "turn_type",
        "facility_type",
        "include_in_peak",
        "include_in_report",
        "count",
        "pcu",
    ]
    if normalized.empty:
        return pd.DataFrame(columns=columns)
    normalized = _with_report_movement_code(normalized)
    keys = columns[:-2]
    return normalized.groupby(keys, dropna=False, as_index=False).agg(count=("count", "sum"), pcu=("pcu", "sum"))


def vehicle_composition(normalized: pd.DataFrame) -> pd.DataFrame:
    if normalized.empty:
        return pd.DataFrame(columns=["vehicle_class", "count", "pcu", "count_share"])
    result = normalized.groupby("vehicle_class", as_index=False).agg(count=("count", "sum"), pcu=("pcu", "sum"))
    total_count = result["count"].sum()
    result["count_share"] = result["count"] / total_count if total_count else 0
    return result


def hourly_movement_pcu(
    normalized: pd.DataFrame,
    mapping: pd.DataFrame,
    movement_code_scheme: str = MOVEMENT_SCHEME_V1,
) -> pd.DataFrame:
    movement_columns = _movement_columns(mapping, movement_code_scheme=movement_code_scheme)
    columns = ["เวลา", *movement_columns, "Total"]
    df = _hourly_frame(normalized)
    if df.empty:
        return pd.DataFrame(columns=columns)
    df = _with_report_movement_code(df)

    pivot = df.pivot_table(index="hour_minute", columns="movement_code", values="pcu", aggfunc="sum", fill_value=0)
    hourly_totals = df.groupby("hour_minute")["pcu"].sum().to_dict()
    for movement in movement_columns:
        if movement not in pivot.columns:
            pivot[movement] = 0
    pivot = pivot[movement_columns].sort_index()
    result = pivot.round(0).astype(int).reset_index()
    result.insert(0, "เวลา", result["hour_minute"].map(_hour_label))
    result["Total"] = result["hour_minute"].map(lambda hour: int(round(hourly_totals.get(hour, 0))))
    result = result.drop(columns=["hour_minute"])

    total_row = {"เวลา": "รวม"}
    for movement in movement_columns:
        total_row[movement] = int(result[movement].sum())
    total_row["Total"] = int(round(df["pcu"].sum()))
    return pd.concat([result, pd.DataFrame([total_row])], ignore_index=True)[columns]


def hourly_vehicle_class(normalized: pd.DataFrame, hourly: pd.DataFrame) -> pd.DataFrame:
    columns = ["เวลา", "Total (PCU)", "Total (คัน)", *VEHICLE_CLASSES]
    df = _hourly_frame(normalized)
    if df.empty:
        return pd.DataFrame(columns=columns)

    counts = df.pivot_table(index="hour_minute", columns="vehicle_class", values="count", aggfunc="sum", fill_value=0)
    for vehicle_class in VEHICLE_CLASSES:
        if vehicle_class not in counts.columns:
            counts[vehicle_class] = 0
    counts = counts[VEHICLE_CLASSES].sort_index()

    pcu_by_hour = hourly.copy()
    if not pcu_by_hour.empty:
        pcu_by_hour["hour_minute"] = pcu_by_hour["hour_start"].map(time_to_minutes)
        pcu_lookup = pcu_by_hour.set_index("hour_minute")["pcu"].to_dict()
    else:
        pcu_lookup = df.groupby("hour_minute")["pcu"].sum().to_dict()

    result = counts.round(0).astype(int).reset_index()
    result.insert(0, "เวลา", result["hour_minute"].map(_hour_label))
    result.insert(1, "Total (PCU)", result["hour_minute"].map(lambda hour: int(round(pcu_lookup.get(hour, 0)))))
    result.insert(2, "Total (คัน)", result[VEHICLE_CLASSES].sum(axis=1))
    result = result.drop(columns=["hour_minute"])

    total_row = {"เวลา": "รวม", "Total (PCU)": int(result["Total (PCU)"].sum()), "Total (คัน)": int(result["Total (คัน)"].sum())}
    for vehicle_class in VEHICLE_CLASSES:
        total_row[vehicle_class] = int(result[vehicle_class].sum())
    return pd.concat([result, pd.DataFrame([total_row])], ignore_index=True)[columns]


def vehicle_composition_report(normalized: pd.DataFrame) -> pd.DataFrame:
    columns = ["ประเภทยานพาหนะ", "vehicle_class", "จำนวนคัน", "สัดส่วน (%)"]
    total_count = normalized["count"].sum() if not normalized.empty and "count" in normalized else 0
    grouped = (
        normalized.groupby("vehicle_class")["count"].sum().to_dict()
        if not normalized.empty and "vehicle_class" in normalized
        else {}
    )

    rows = []
    for vehicle_class in VEHICLE_CLASSES:
        count = grouped.get(vehicle_class, 0)
        rows.append(
            {
                "ประเภทยานพาหนะ": THAI_VEHICLE_LABELS[vehicle_class],
                "vehicle_class": vehicle_class,
                "จำนวนคัน": int(round(count)),
                "สัดส่วน (%)": count / total_count if total_count else 0,
            }
        )
    rows.append(
        {
            "ประเภทยานพาหนะ": "รวม",
            "vehicle_class": "Total",
            "จำนวนคัน": int(round(total_count)),
            "สัดส่วน (%)": 1 if total_count else 0,
        }
    )
    return pd.DataFrame(rows, columns=columns)


def vehicle_group_pce(normalized: pd.DataFrame) -> pd.DataFrame:
    total_count = normalized["count"].sum() if not normalized.empty and "count" in normalized else 0
    grouped = (
        normalized.groupby("vehicle_class")["count"].sum().to_dict()
        if not normalized.empty and "vehicle_class" in normalized
        else {}
    )
    rows = []
    for group in DEFAULT_VEHICLE_GROUPS:
        value = sum(grouped.get(vehicle_class, 0) for vehicle_class in group["classes"])
        rows.append(
            {
                "Type": group["Type"],
                "PCE": group["PCE"],
                "value": int(round(value)),
                "percent": value / total_count if total_count else 0,
                "classes": ", ".join(group["classes"]),
            }
        )
    return pd.DataFrame(rows, columns=["Type", "PCE", "value", "percent", "classes"])


def phf_15min(normalized: pd.DataFrame) -> pd.DataFrame:
    columns = ["เวลา", "15", "30", "45", "60", "Total", "PHF"]
    if normalized.empty:
        return pd.DataFrame(columns=columns)

    df = normalized[normalized["include_in_peak"]].copy() if "include_in_peak" in normalized else normalized.copy()
    df["minute"] = df["time_start"].map(time_to_minutes)
    df = df.dropna(subset=["minute"])
    if df.empty:
        return pd.DataFrame(columns=columns)

    df["hour_minute"] = (df["minute"].astype(int) // 60) * 60
    df["quarter"] = ((df["minute"].astype(int) - df["hour_minute"]) // 15).astype(int)
    grouped = df.groupby(["hour_minute", "quarter"], as_index=False).agg(pcu=("pcu", "sum"))

    rows = []
    for hour_minute in sorted(grouped["hour_minute"].unique()):
        hour = grouped[grouped["hour_minute"] == hour_minute]
        values = [float(hour.loc[hour["quarter"] == quarter, "pcu"].sum()) for quarter in range(4)]
        total = sum(values)
        max_15 = max(values) if values else 0
        rows.append(
            {
                "เวลา": _hour_label(int(hour_minute)),
                "15": round(values[0]),
                "30": round(values[1]),
                "45": round(values[2]),
                "60": round(values[3]),
                "Total": round(total),
                "PHF": total / (4 * max_15) if max_15 else 0,
            }
        )
    return pd.DataFrame(rows, columns=columns)


def movement_aggregation_audit(normalized: pd.DataFrame, mapping: pd.DataFrame) -> pd.DataFrame:
    columns = [
        "output_movement_code",
        "source_direction",
        "source_stream",
        "raw_movement_label",
        "include_in_report",
        "include_in_peak",
        "total_vehicle_count",
        "total_pcu",
    ]
    if mapping.empty:
        return pd.DataFrame(columns=columns)

    mapping = mapping.copy()
    if "output_movement_code" not in mapping.columns:
        mapping["output_movement_code"] = mapping["movement_code"] if "movement_code" in mapping else ""
    if "source_direction" not in mapping.columns:
        mapping["source_direction"] = mapping["raw_direction"] if "raw_direction" in mapping else ""
    if "source_stream" not in mapping.columns:
        mapping["source_stream"] = "mainline"
    if "raw_movement_label" not in mapping.columns:
        mapping["raw_movement_label"] = mapping["source_direction"]
    for column, default in {"include_in_report": True, "include_in_peak": True}.items():
        if column not in mapping.columns:
            mapping[column] = default

    keys = ["output_movement_code", "source_direction", "source_stream", "raw_movement_label", "include_in_report", "include_in_peak"]
    audit = mapping[keys].copy()
    audit["output_movement_code"] = audit["output_movement_code"].fillna("").astype(str)
    audit["source_direction"] = audit["source_direction"].fillna("").astype(str)
    audit["source_stream"] = audit["source_stream"].fillna("").astype(str).str.strip().replace("", "mainline")
    audit["raw_movement_label"] = audit["raw_movement_label"].fillna("").astype(str)

    if not normalized.empty:
        df = normalized.copy()
        if "output_movement_code" not in df.columns:
            df["output_movement_code"] = df["movement_code"] if "movement_code" in df else ""
        if "source_direction" not in df.columns:
            df["source_direction"] = df["raw_direction"] if "raw_direction" in df else ""
        if "source_stream" not in df.columns:
            df["source_stream"] = "mainline"
        if "raw_movement_label" not in df.columns:
            df["raw_movement_label"] = df["source_direction"]
        df["source_stream"] = df["source_stream"].fillna("").astype(str).str.strip().replace("", "mainline")
        totals = (
            df.groupby(["output_movement_code", "source_direction", "source_stream", "raw_movement_label"], dropna=False, as_index=False)
            .agg(total_vehicle_count=("count", "sum"), total_pcu=("pcu", "sum"))
        )
        audit = audit.merge(
            totals,
            on=["output_movement_code", "source_direction", "source_stream", "raw_movement_label"],
            how="left",
        )
    else:
        audit["total_vehicle_count"] = 0
        audit["total_pcu"] = 0

    audit["total_vehicle_count"] = audit["total_vehicle_count"].fillna(0)
    audit["total_pcu"] = audit["total_pcu"].fillna(0)
    return audit[columns]
