"""Project paths for local runs and Google Colab."""

from __future__ import annotations

import os
from pathlib import Path


def in_colab() -> bool:
    try:
        import google.colab  # noqa: F401

        return True
    except ImportError:
        return False


CODE_DIR = Path(__file__).resolve().parent

DEFAULT_COLAB_DATA_ROOT = Path("/content/drive/Shareddrives/FA Ops Europe: Rate Maintenance Team /Documents/AI Adoption RMT/RMT_Puma/RMT_Update")

# Alternate layout without trailing space on the shared-drive folder name.
LEGACY_COLAB_DATA_ROOT = Path("/content/drive/Shareddrives/FA Ops Europe: Rate Maintenance Team /Documents/AI Adoption RMT/RMT_Puma/RMT_Update")

PREVIOUS_RA_SUBDIR = "previous RA"
UPDATE_SUBDIR = "update"

INPUT_SUBDIR = "input"
OUTPUT_SUBDIR = "output"
PROCESSING_SUBDIR = "processing"


def default_colab_data_root() -> Path:
    if DEFAULT_COLAB_DATA_ROOT.exists():
        return DEFAULT_COLAB_DATA_ROOT
    if LEGACY_COLAB_DATA_ROOT.exists():
        return LEGACY_COLAB_DATA_ROOT
    return DEFAULT_COLAB_DATA_ROOT


def data_root() -> Path:
    override = os.environ.get("PUMA_UPDATE_DATA_ROOT", "").strip()
    if override:
        return Path(override)
    if in_colab():
        return default_colab_data_root()
    return CODE_DIR


DATA_ROOT = data_root()
INPUT_DIR = DATA_ROOT / INPUT_SUBDIR
PREVIOUS_RA_DIR = INPUT_DIR / PREVIOUS_RA_SUBDIR
UPDATE_DIR = INPUT_DIR / UPDATE_SUBDIR
OUTPUT_DIR = DATA_ROOT / OUTPUT_SUBDIR
PROCESSING_DIR = DATA_ROOT / PROCESSING_SUBDIR


def configure_colab_drive(*, force_remount: bool = False) -> None:
    """Mount Google Drive when running in Colab. No-op locally."""
    if not in_colab():
        return
    from google.colab import drive

    drive.mount("/content/drive", force_remount=force_remount)


def ensure_data_dirs() -> None:
    for path in (INPUT_DIR, PREVIOUS_RA_DIR, UPDATE_DIR, OUTPUT_DIR, PROCESSING_DIR):
        path.mkdir(parents=True, exist_ok=True)


def print_paths() -> None:
    label = "Google Colab" if in_colab() else "local"
    print(f"Environment: {label}")
    print(f"  Code:         {CODE_DIR}")
    print(f"  Data root:    {DATA_ROOT}")
    print(f"  Previous RA:  {PREVIOUS_RA_DIR}")
    print(f"  Update:       {UPDATE_DIR}")
    print(f"  Processing:   {PROCESSING_DIR}")
    print(f"  Output:       {OUTPUT_DIR}")
