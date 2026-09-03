"""Google Colab entry point for the Heppner pipeline."""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

CODE_DIR = Path("/content/Corning-Heppner")
REQUIREMENTS = CODE_DIR / "requirements.txt"


def install_dependencies() -> None:
    if REQUIREMENTS.exists():
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", "-q", "-r", str(REQUIREMENTS)]
        )
    else:
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", "-q", "pandas", "openpyxl"]
        )


def main() -> None:
    if CODE_DIR.exists() and str(CODE_DIR) not in sys.path:
        sys.path.insert(0, str(CODE_DIR))

    install_dependencies()

    from paths import exit_with_code
    from run_pipeline import main as run_main

    exit_with_code(run_main())


if __name__ == "__main__":
    main()
