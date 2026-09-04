"""Resolve lane identity columns across Sea and Air rate layouts."""

from __future__ import annotations

import pandas as pd

from load_input import destination_country_code

ORIGIN_COUNTRY_COLUMN = "ORIGIN COUNTRY"
DESTINATION_COUNTRY_COLUMN = "DESTINATION COUNTRY"

ORIGIN_COUNTRY_ALIASES = (ORIGIN_COUNTRY_COLUMN,)
DESTINATION_COUNTRY_ALIASES = (DESTINATION_COUNTRY_COLUMN,)

ORIGIN_PORT_ALIASES = (
    "ORIGIN PORT",
    "Origin Port",
    "ORIGIN AIRPORT",
    "ORIGIN AIRPORT NAME",
)

DESTINATION_PORT_ALIASES = (
    "DESTINATION PORT",
    "Destination Port",
    "DESTINATION AIRPORT",
    "DESTINATION AIRPORT NAME",
)

LANE_ID_ALIASES = ("ID", "Lane ID")
LANE_NUMBER_ALIASES = ("Lane #",)
VALID_FROM_ALIASES = ("Valid from", "VALID FROM", "RATE VALID FROM")
VALID_TO_ALIASES = ("Valid to", "VALID TO", "RATE VALID TO")

LANE_KEY_COLUMN = "_lane_key"


def normalize_port_name(value: object) -> str | None:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    text = str(value).strip().upper()
    if not text:
        return None
    # Istanbul airports are represented as IST/SAW in RA and IST in update files.
    if text in {"IST", "SAW", "IST/SAW"}:
        return "IST/SAW"
    return text


def build_lane_key(
    origin_country: object,
    destination_country: object,
    origin_port: object,
    destination_port: object,
) -> str | None:
    origin_cc = destination_country_code(origin_country)
    dest_cc = destination_country_code(destination_country)
    origin_port_name = normalize_port_name(origin_port)
    dest_port_name = normalize_port_name(destination_port)
    if not origin_cc or not dest_cc or not origin_port_name or not dest_port_name:
        return None
    return f"{origin_cc}|{dest_cc}|{origin_port_name}|{dest_port_name}"


def _column_lookup(columns: pd.Index | list[str]) -> dict[str, str]:
    return {str(column).strip().upper(): str(column) for column in columns}


def find_column(columns: pd.Index | list[str], aliases: tuple[str, ...]) -> str | None:
    lookup = _column_lookup(columns)
    for alias in aliases:
        match = lookup.get(alias.strip().upper())
        if match is not None:
            return match
    return None


def find_header_column(column_map: dict[str, int], aliases: tuple[str, ...]) -> str | None:
    lookup = {str(name).strip().upper(): name for name in column_map}
    for alias in aliases:
        match = lookup.get(alias.strip().upper())
        if match is not None:
            return match
    return None


def resolve_lane_columns(columns: pd.Index | list[str]) -> dict[str, str]:
    resolved: dict[str, str] = {}
    spec = (
        ("origin_country", ORIGIN_COUNTRY_ALIASES),
        ("destination_country", DESTINATION_COUNTRY_ALIASES),
        ("origin_port", ORIGIN_PORT_ALIASES),
        ("destination_port", DESTINATION_PORT_ALIASES),
    )
    missing: list[str] = []
    for key, aliases in spec:
        column = find_column(columns, aliases)
        if column is None:
            missing.append(" / ".join(aliases))
        else:
            resolved[key] = column
    if missing:
        raise ValueError(f"Missing lane columns: {', '.join(missing)}")
    return resolved


def resolve_header_lane_columns(column_map: dict[str, int]) -> dict[str, str]:
    resolved: dict[str, str] = {}
    spec = (
        ("origin_country", ORIGIN_COUNTRY_ALIASES),
        ("destination_country", DESTINATION_COUNTRY_ALIASES),
        ("origin_port", ORIGIN_PORT_ALIASES),
        ("destination_port", DESTINATION_PORT_ALIASES),
    )
    missing: list[str] = []
    for key, aliases in spec:
        column = find_header_column(column_map, aliases)
        if column is None:
            missing.append(" / ".join(aliases))
        else:
            resolved[key] = column
    if missing:
        raise ValueError(f"Missing lane columns: {', '.join(missing)}")
    return resolved


def lane_key_from_row(row: pd.Series, lane_columns: dict[str, str]) -> str | None:
    return build_lane_key(
        row[lane_columns["origin_country"]],
        row[lane_columns["destination_country"]],
        row[lane_columns["origin_port"]],
        row[lane_columns["destination_port"]],
    )


def lane_key_from_values(
    values: list[object],
    column_map: dict[str, int],
    lane_columns: dict[str, str],
) -> str | None:
    return build_lane_key(
        values[column_map[lane_columns["origin_country"]] - 1],
        values[column_map[lane_columns["destination_country"]] - 1],
        values[column_map[lane_columns["origin_port"]] - 1],
        values[column_map[lane_columns["destination_port"]] - 1],
    )


def lane_key_description(lane_columns: dict[str, str]) -> str:
    return (
        f"{lane_columns['origin_country']} + {lane_columns['destination_country']} + "
        f"{lane_columns['origin_port']} + {lane_columns['destination_port']}"
    )


def resolve_rate_card_identity_columns(
    column_map: dict[str, int],
    *,
    require_validity: bool = True,
) -> dict[str, str]:
    resolved: dict[str, str] = {}
    required = (
        ("lane_id", LANE_ID_ALIASES),
        ("valid_from", VALID_FROM_ALIASES),
        ("valid_to", VALID_TO_ALIASES),
    )
    missing: list[str] = []
    for key, aliases in required:
        column = find_header_column(column_map, aliases)
        if column is None:
            if key == "lane_id" or require_validity:
                missing.append(" / ".join(aliases))
        else:
            resolved[key] = column

    lane_number = find_header_column(column_map, LANE_NUMBER_ALIASES)
    if lane_number is not None:
        resolved["lane_number"] = lane_number

    if missing:
        raise ValueError(f"Rate card is missing columns: {', '.join(missing)}")
    return resolved


def add_lane_key(df: pd.DataFrame) -> pd.DataFrame:
    lane_columns = resolve_lane_columns(df.columns)
    out = df.copy()
    out[LANE_KEY_COLUMN] = out.apply(lambda row: lane_key_from_row(row, lane_columns), axis=1)
    return out
