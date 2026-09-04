"""Pick input files from previous RA and update folders, load to DataFrames, save xlsx to processing/."""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

import pandas as pd

from paths import PREVIOUS_RA_DIR, PROCESSING_DIR, UPDATE_DIR

SKIP_SHEETS = {"selections"}
SUPPORTED_SUFFIXES = {".xlsx", ".xlsm", ".xls", ".csv"}

ROLE_PREVIOUS_RA = "previous_ra"
ROLE_UPDATE = "update"

PREVIOUS_RA_SHEET = "rate card"
UPDATE_DESTINATION_COLUMN = "DESTINATION COUNTRY"
COUNTRY_CODE_RE = re.compile(r"\(([A-Z]{2})\)\s*$")

SKIP_ANSWERS = {"", "skip", "none", "n"}


INVALID_SHEET_NAME = re.compile(r'[\[\]:*?/\\]')
MAX_SHEET_NAME_LEN = 31


@dataclass
class FileSelection:
    previous_ra: list[Path]
    update: Path

    def labeled_paths(self) -> list[tuple[str, Path]]:
        items = [(ROLE_PREVIOUS_RA, path) for path in self.previous_ra]
        items.append((ROLE_UPDATE, self.update))
        return items


@dataclass
class LoadedData:
    previous_ra: dict[str, pd.DataFrame]
    update: dict[str, pd.DataFrame]


def _unique_sheet_name(stem: str, used: set[str]) -> str:
    base = INVALID_SHEET_NAME.sub("_", stem).strip(" '")
    base = base[:MAX_SHEET_NAME_LEN] or "sheet"
    name = base
    counter = 2
    while name.casefold() in {existing.casefold() for existing in used}:
        suffix = f"_{counter}"
        name = f"{base[: MAX_SHEET_NAME_LEN - len(suffix)]}{suffix}"
        counter += 1
    used.add(name)
    return name


def destination_country_code(value: object) -> str | None:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    text = str(value).strip().upper()
    if re.fullmatch(r"[A-Z]{2}", text):
        return text
    match = COUNTRY_CODE_RE.search(str(value).strip())
    return match.group(1) if match else None


def _find_destination_country_column(df: pd.DataFrame) -> str:
    for column in df.columns:
        if str(column).strip().upper() == UPDATE_DESTINATION_COLUMN:
            return str(column)
    raise ValueError(f"Column '{UPDATE_DESTINATION_COLUMN}' not found in previous RA data.")


def collect_previous_ra_destination_codes(previous_ra: dict[str, pd.DataFrame]) -> set[str]:
    codes: set[str] = set()
    for frame in previous_ra.values():
        column = _find_destination_country_column(frame)
        for value in frame[column].dropna().unique():
            code = destination_country_code(value)
            if code:
                codes.add(code)
    if not codes:
        raise ValueError(f"No destination country codes found in previous RA '{UPDATE_DESTINATION_COLUMN}' values.")
    return codes


def filter_update_sheets_by_countries(
    update_sheets: dict[str, pd.DataFrame],
    country_codes: set[str],
) -> dict[str, pd.DataFrame]:
    allowed = {code.upper() for code in country_codes}
    filtered = {
        sheet_name: frame
        for sheet_name, frame in update_sheets.items()
        if sheet_name.upper() in allowed
    }

    removed = sorted(set(update_sheets) - set(filtered))
    missing = sorted(allowed - {sheet_name.upper() for sheet_name in filtered})
    if removed:
        print(f"  Removed update tabs not in previous RA: {', '.join(removed)}")
    if missing:
        print(f"  Warning: previous RA countries without update tabs: {', '.join(missing)}")
    if not filtered:
        raise ValueError("No update tabs matched previous RA destination countries.")
    return filtered


def _destination_country_sheet_name(value: object, used: set[str]) -> str:
    text = "" if pd.isna(value) else str(value).strip()
    match = re.search(r"\(([^)]+)\)\s*$", text)
    stem = match.group(1) if match else text
    return _unique_sheet_name(stem or "unknown", used)


def split_update_by_destination_country(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    if UPDATE_DESTINATION_COLUMN not in df.columns:
        raise ValueError(f"Column '{UPDATE_DESTINATION_COLUMN}' not found in update data.")

    sheets: dict[str, pd.DataFrame] = {}
    used_sheet_names: set[str] = set()
    for country, group in df.groupby(UPDATE_DESTINATION_COLUMN, dropna=False, sort=True):
        sheet_name = _destination_country_sheet_name(country, used_sheet_names)
        sheets[sheet_name] = group.reset_index(drop=True)
    return sheets


def list_input_files(input_dir: Path) -> list[Path]:
    if not input_dir.exists():
        raise FileNotFoundError(f"Input folder not found: {input_dir}")

    files = [
        path
        for path in sorted(input_dir.iterdir())
        if path.is_file()
        and path.suffix.lower() in SUPPORTED_SUFFIXES
        and not path.name.startswith("~$")
    ]
    if not files:
        raise FileNotFoundError(f"No supported files found in {input_dir}")
    return files


def print_file_list(files: list[Path], *, folder_label: str) -> None:
    print(f"\nFiles in {folder_label}:")
    for index, path in enumerate(files, start=1):
        print(f"  {index}. {path.name}")


def _parse_tokens(raw: str, available: list[int]) -> list[int]:
    available_set = set(available)
    display_numbers = {index + 1 for index in available}
    tokens = [token for token in re.split(r"[,\s]+", raw) if token]
    if not tokens:
        raise ValueError("Please enter numbers from the list, a range like 1-3, all, or skip.")

    indices: list[int] = []
    for token in tokens:
        if re.fullmatch(r"\d+-\d+", token):
            start_text, end_text = token.split("-", 1)
            start, end = int(start_text), int(end_text)
            if start > end:
                raise ValueError(f"Range {token} is invalid.")
            for number in range(start, end + 1):
                if number not in display_numbers:
                    raise ValueError(f"{number} is not available.")
                indices.append(number - 1)
            continue
        if not token.isdigit():
            raise ValueError(f"'{token}' is not a valid choice. Use numbers, ranges, all, or skip.")
        number = int(token)
        if number - 1 not in available_set:
            raise ValueError(f"{number} is not available.")
        indices.append(number - 1)

    unique: list[int] = []
    seen: set[int] = set()
    for index in indices:
        if index not in seen:
            unique.append(index)
            seen.add(index)
    return unique


def parse_choice_indices(
    raw: str,
    available: list[int],
    *,
    allow_empty: bool,
    allow_multiple: bool,
) -> list[int]:
    text = raw.strip()
    if text.lower() in SKIP_ANSWERS:
        if allow_empty:
            return []
        raise ValueError("This choice cannot be skipped.")

    if text.lower() == "all":
        if not available:
            raise ValueError("No files are left to choose.")
        if not allow_multiple and len(available) > 1:
            raise ValueError("Please choose a single file.")
        return list(available)

    indices = _parse_tokens(text, available)
    if not allow_multiple and len(indices) != 1:
        raise ValueError("Please choose a single file.")
    if allow_multiple and not indices:
        raise ValueError("Please choose at least one file.")
    return indices


def _available_hint(available: list[int], *, allow_empty: bool, allow_multiple: bool) -> str:
    if not available:
        return "skip"
    numbers = ", ".join(str(index + 1) for index in available)
    parts = [numbers]
    if allow_multiple and len(available) > 1:
        parts.append("e.g. 1 2, 1-3, or all")
    if allow_empty:
        parts.append("or skip")
    return "; ".join(parts)


def prompt_indices(
    *,
    prompt: str,
    available: list[int],
    allow_empty: bool = False,
    allow_multiple: bool = False,
) -> list[int]:
    hint = _available_hint(available, allow_empty=allow_empty, allow_multiple=allow_multiple)
    while True:
        raw = input(f"\n{prompt} [{hint}]: ").strip()
        try:
            return parse_choice_indices(
                raw,
                available,
                allow_empty=allow_empty,
                allow_multiple=allow_multiple,
            )
        except ValueError as exc:
            print(exc)


def prompt_file_selection() -> FileSelection:
    previous_ra_files = list_input_files(PREVIOUS_RA_DIR)
    update_files = list_input_files(UPDATE_DIR)

    print_file_list(previous_ra_files, folder_label="previous RA")
    previous_ra_indices = prompt_indices(
        prompt="Choose previous RA file(s)",
        available=list(range(len(previous_ra_files))),
        allow_multiple=True,
    )

    print_file_list(update_files, folder_label="update")
    update_indices = prompt_indices(
        prompt="Choose update file",
        available=list(range(len(update_files))),
    )

    selection = FileSelection(
        previous_ra=[previous_ra_files[index] for index in previous_ra_indices],
        update=update_files[update_indices[0]],
    )
    print("\nSelected files:")
    for role, path in selection.labeled_paths():
        print(f"  {role}: {path.name}")
    return selection


def _find_sheet_name(sheet_names: list[str], target: str) -> str | None:
    normalized = target.strip().lower()
    for name in sheet_names:
        if name.strip().lower() == normalized:
            return name
    return None


def _find_rate_card_header_row(path: Path, sheet_name: str) -> int:
    preview = pd.read_excel(path, sheet_name=sheet_name, header=None, nrows=40)
    for row_index in range(len(preview)):
        normalized = {
            str(value).strip().upper()
            for value in preview.iloc[row_index].tolist()
            if pd.notna(value) and str(value).strip()
        }
        if "LANE #" in normalized and UPDATE_DESTINATION_COLUMN in normalized:
            return row_index
    raise ValueError(f"Could not find Rate card header row with '{UPDATE_DESTINATION_COLUMN}' in {path.name}")


def load_previous_ra_dataframe(path: Path) -> pd.DataFrame:
    suffix = path.suffix.lower()
    if suffix == ".csv":
        raise ValueError(f"Previous RA must be an Excel file with a '{PREVIOUS_RA_SHEET}' sheet: {path.name}")
    if suffix not in {".xlsx", ".xlsm", ".xls"}:
        raise ValueError(f"Unsupported file type: {path.name}")

    workbook = pd.ExcelFile(path)
    sheet_name = _find_sheet_name(workbook.sheet_names, PREVIOUS_RA_SHEET)
    if sheet_name is None:
        available = ", ".join(workbook.sheet_names)
        raise ValueError(
            f"Sheet '{PREVIOUS_RA_SHEET}' not found in {path.name}. Available sheets: {available}"
        )

    header_row = _find_rate_card_header_row(path, sheet_name)
    frame = pd.read_excel(workbook, sheet_name=sheet_name, header=header_row)
    if frame.empty:
        raise ValueError(f"Sheet '{sheet_name}' is empty in {path.name}")
    return frame


def _read_excel_sheets(path: Path) -> dict[str, pd.DataFrame]:
    workbook = pd.ExcelFile(path)
    frames: dict[str, pd.DataFrame] = {}
    for sheet_name in workbook.sheet_names:
        if sheet_name.strip().lower() in SKIP_SHEETS:
            continue
        frame = pd.read_excel(workbook, sheet_name=sheet_name)
        if frame.empty:
            continue
        frames[sheet_name] = frame
    return frames


def load_update_dataframe(path: Path) -> pd.DataFrame:
    suffix = path.suffix.lower()
    if suffix == ".csv":
        return pd.read_csv(path)

    if suffix not in {".xlsx", ".xlsm", ".xls"}:
        raise ValueError(f"Unsupported file type: {path.name}")

    sheets = _read_excel_sheets(path)
    if not sheets:
        raise ValueError(f"No data sheets found in {path.name}")

    if len(sheets) == 1:
        return next(iter(sheets.values()))

    combined = []
    for sheet_name, frame in sheets.items():
        copy = frame.copy()
        copy.insert(0, "_sheet", sheet_name)
        combined.append(copy)
    return pd.concat(combined, ignore_index=True, sort=False)


def load_selection(selection: FileSelection) -> LoadedData:
    previous_ra_sheets: dict[str, pd.DataFrame] = {}
    used_sheet_names: set[str] = set()

    for path in selection.previous_ra:
        sheet_name = _unique_sheet_name(path.stem, used_sheet_names)
        print(f"Loading {ROLE_PREVIOUS_RA}: {path.name} (sheet: Rate card) ...")
        frame = load_previous_ra_dataframe(path).copy()
        frame.insert(0, "_source_file", path.name)
        frame.insert(0, "_file_role", ROLE_PREVIOUS_RA)
        previous_ra_sheets[sheet_name] = frame

    print(f"Loading {ROLE_UPDATE}: {selection.update.name} ...")
    update = load_update_dataframe(selection.update).copy()
    update.insert(0, "_source_file", selection.update.name)
    update.insert(0, "_file_role", ROLE_UPDATE)
    update_sheets = split_update_by_destination_country(update)

    destination_codes = collect_previous_ra_destination_codes(previous_ra_sheets)
    code_list = ", ".join(sorted(destination_codes))
    print(f"Previous RA destination countries: {code_list}")
    print("Filtering update tabs to matching destination countries ...")
    update_sheets = filter_update_sheets_by_countries(update_sheets, destination_codes)

    return LoadedData(previous_ra=previous_ra_sheets, update=update_sheets)


def save_workbook(sheets: dict[str, pd.DataFrame], role: str, output_dir: Path = PROCESSING_DIR) -> Path:
    if not sheets:
        raise ValueError(f"No sheets to save for {role}.")

    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / f"{role}.xlsx"
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, frame in sheets.items():
            frame.to_excel(writer, sheet_name=sheet_name, index=False)
    return output_path


def run_load_input(selection: FileSelection | None = None) -> tuple[dict[str, Path], FileSelection]:
    if selection is None:
        selection = prompt_file_selection()

    loaded = load_selection(selection)
    saved: dict[str, Path] = {}

    previous_ra_path = save_workbook(loaded.previous_ra, ROLE_PREVIOUS_RA)
    saved[ROLE_PREVIOUS_RA] = previous_ra_path
    total_rows = sum(len(frame) for frame in loaded.previous_ra.values())
    tab_names = ", ".join(loaded.previous_ra)
    print(f"Saved {ROLE_PREVIOUS_RA} ({total_rows:,} rows, {len(loaded.previous_ra)} tab(s): {tab_names}) to {previous_ra_path.name}")

    update_path = save_workbook(loaded.update, ROLE_UPDATE)
    saved[ROLE_UPDATE] = update_path
    update_rows = sum(len(frame) for frame in loaded.update.values())
    update_tabs = ", ".join(loaded.update)
    print(
        f"Saved {ROLE_UPDATE} ({update_rows:,} rows, {len(loaded.update)} tab(s) by "
        f"{UPDATE_DESTINATION_COLUMN}: {update_tabs}) to {update_path.name}"
    )
    return saved, selection


def main() -> FileSelection:
    from paths import configure_colab_drive, ensure_data_dirs, print_paths

    configure_colab_drive()
    ensure_data_dirs()
    print_paths()
    print()
    _, selection = run_load_input()
    return selection


if __name__ == "__main__":
    main()
