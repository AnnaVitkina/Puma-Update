"""Run the full Puma rate-card update pipeline (local machine or Google Colab).

Google Colab
------------
1. Upload or clone this folder to ``/content/Puma-Update``.
2. Put data on Google Drive under the shared RMT_Update folder (see ``paths.py``).
3. Run::

       import os
       exec(open("/content/Puma-Update/run_pipeline.py").read())

   Optional: set a custom data folder before running::

       import os
       os.environ["PUMA_UPDATE_DATA_ROOT"] = "/content/drive/Shareddrives/.../RMT_Update"
       exec(open("/content/Puma-Update/run_pipeline.py").read())

Local machine
-------------
::

    python run_pipeline.py

Or::

    python run_load_input.py
"""

from __future__ import annotations

import sys
from pathlib import Path


def code_dir() -> Path:
    """Resolve the folder that contains the project Python modules."""
    file_path = globals().get("__file__")
    if file_path:
        return Path(file_path).resolve().parent

    colab_dir = Path("/content/Puma-Update")
    if colab_dir.is_dir():
        return colab_dir

    return Path.cwd()


def bootstrap() -> Path:
    """Add the project folder to ``sys.path`` so imports work after ``exec()``."""
    project_dir = code_dir()
    project_str = str(project_dir)
    if project_str not in sys.path:
        sys.path.insert(0, project_str)
    return project_dir


def main() -> None:
    bootstrap()

    from load_input import main as load_main
    from map_lanes import run_map_lanes
    from paths import in_colab
    from rate_card_update import run_apply_rate_card_update

    selection = load_main()
    if in_colab():
        print("\nTip: input files live on Drive under input/previous RA and input/update.")
    run_map_lanes()
    run_apply_rate_card_update(previous_ra_paths=selection.previous_ra)


if __name__ == "__main__":
    main()
