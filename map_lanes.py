"""Map update lanes to previous RA lanes by geography; export unmapped update lanes."""

from __future__ import annotations

from pathlib import Path

import pandas as pd

from lane_columns import (
    LANE_KEY_COLUMN,
    ORIGIN_COUNTRY_COLUMN,
    DESTINATION_COUNTRY_COLUMN,
    add_lane_key,
    lane_key_description,
    resolve_lane_columns,
)
from load_input import ROLE_PREVIOUS_RA, ROLE_UPDATE
from paths import OUTPUT_DIR, PROCESSING_DIR

UNMAPPED_OUTPUT_NAME = "unmapped_lanes.xlsx"

BASE_UPDATE_REFERENCE_COLUMNS = (
    "LANE ID",
    "_update_tab",
    "_source_file",
    ORIGIN_COUNTRY_COLUMN,
    DESTINATION_COUNTRY_COLUMN,
    "MODE",
    "SERVICE TYPE",
    "CORE CARRIER",
    "CARRIER PORTFOLIO",
)


def load_processing_workbook(path: Path) -> dict[str, pd.DataFrame]:
    if not path.exists():
        raise FileNotFoundError(f"Processing file not found: {path}")
    workbook = pd.ExcelFile(path)
    return {sheet_name: pd.read_excel(workbook, sheet_name=sheet_name) for sheet_name in workbook.sheet_names}


def combine_workbook_sheets(sheets: dict[str, pd.DataFrame], tab_column: str) -> pd.DataFrame:
    if not sheets:
        raise ValueError("No sheets to combine.")
    frames: list[pd.DataFrame] = []
    for tab_name, frame in sheets.items():
        copy = frame.copy()
        copy.insert(0, tab_column, tab_name)
        frames.append(copy)
    return pd.concat(frames, ignore_index=True, sort=False)


def _select_update_columns(df: pd.DataFrame, lane_columns: dict[str, str]) -> pd.DataFrame:
    selected = [column for column in BASE_UPDATE_REFERENCE_COLUMNS if column in df.columns]
    for key in ("origin_port", "destination_port"):
        column = lane_columns[key]
        if column not in selected:
            selected.append(column)
    extra = [LANE_KEY_COLUMN, "_unmapped_reason"] if LANE_KEY_COLUMN in df.columns else []
    return df[selected + extra].copy()


def _mark_update_unmapped_reason(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out["_unmapped_reason"] = out[LANE_KEY_COLUMN].apply(
        lambda value: "missing_lane_key_fields"
        if value is None or (isinstance(value, float) and pd.isna(value))
        else "no_matching_previous_ra_lane"
    )
    return out


def map_lanes(
    previous_ra: dict[str, pd.DataFrame] | pd.DataFrame,
    update: dict[str, pd.DataFrame] | pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    if isinstance(previous_ra, dict):
        ra = combine_workbook_sheets(previous_ra, "_ra_tab")
    else:
        ra = previous_ra.copy()

    if isinstance(update, dict):
        up = combine_workbook_sheets(update, "_update_tab")
    else:
        up = update.copy()

    ra_lane_columns = resolve_lane_columns(ra.columns)
    update_lane_columns = resolve_lane_columns(up.columns)

    ra = add_lane_key(ra)
    up = add_lane_key(up)

    ra_keys = set(ra[LANE_KEY_COLUMN].dropna())
    matched_keys = ra_keys & set(up[LANE_KEY_COLUMN].dropna())
    update_unmapped = _mark_update_unmapped_reason(up.loc[~up[LANE_KEY_COLUMN].isin(ra_keys)].copy())

    summary = pd.DataFrame(
        [
            {"metric": "update_rows", "value": len(up)},
            {"metric": "matched_update_rows", "value": int(up[LANE_KEY_COLUMN].isin(ra_keys).sum())},
            {"metric": "matched_lane_keys", "value": len(matched_keys)},
            {"metric": "update_unmapped_rows", "value": len(update_unmapped)},
            {"metric": "update_missing_lane_key_rows", "value": int(up[LANE_KEY_COLUMN].isna().sum())},
            {"metric": "ra_lane_key", "value": lane_key_description(ra_lane_columns)},
            {"metric": "update_lane_key", "value": lane_key_description(update_lane_columns)},
        ]
    )
    return _select_update_columns(update_unmapped, update_lane_columns), summary


def save_unmapped_lanes(
    update_unmapped: pd.DataFrame,
    summary: pd.DataFrame,
    output_path: Path | None = None,
) -> Path:
    if output_path is None:
        output_path = OUTPUT_DIR / UNMAPPED_OUTPUT_NAME
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="summary", index=False)
        update_unmapped.to_excel(writer, sheet_name="update_unmapped", index=False)
    return output_path


def run_map_lanes(
    *,
    previous_ra_path: Path | None = None,
    update_path: Path | None = None,
    output_path: Path | None = None,
) -> Path:
    previous_ra_path = previous_ra_path or (PROCESSING_DIR / f"{ROLE_PREVIOUS_RA}.xlsx")
    update_path = update_path or (PROCESSING_DIR / f"{ROLE_UPDATE}.xlsx")

    print("\n--- Map update lanes to previous RA ---")
    print(f"  Previous RA: {previous_ra_path.name}")
    print(f"  Update:      {update_path.name}")

    previous_ra = load_processing_workbook(previous_ra_path)
    update = load_processing_workbook(update_path)
    ra_lane_columns = resolve_lane_columns(combine_workbook_sheets(previous_ra, "_ra_tab").columns)
    update_lane_columns = resolve_lane_columns(combine_workbook_sheets(update, "_update_tab").columns)
    print(f"  RA match key:     {lane_key_description(ra_lane_columns)}")
    print(f"  Update match key: {lane_key_description(update_lane_columns)}")

    update_unmapped, summary = map_lanes(previous_ra, update)
    saved_path = save_unmapped_lanes(update_unmapped, summary, output_path)

    matched_rows = int(summary.loc[summary["metric"] == "matched_update_rows", "value"].iloc[0])
    update_rows = int(summary.loc[summary["metric"] == "update_rows", "value"].iloc[0])
    print(f"  Matched update rows: {matched_rows:,} / {update_rows:,}")
    print(f"  Update unmapped rows: {len(update_unmapped):,}")
    print(f"  Saved unmapped update lanes to {saved_path}")
    return saved_path


def main() -> None:
    from paths import configure_colab_drive, ensure_data_dirs, print_paths

    configure_colab_drive()
    ensure_data_dirs()
    print_paths()
    run_map_lanes()


if __name__ == "__main__":
    main()
