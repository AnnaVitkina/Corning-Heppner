"""Shared path configuration for local and Google Colab execution."""

from __future__ import annotations

import os
import sys
from dataclasses import dataclass
from pathlib import Path

COLAB_CODE_DIR = Path("/content/Corning-Heppner")
COLAB_DATA_ROOT = Path(
    "/content/drive/Shareddrives/FA Ops Europe: Rate Maintenance Team "
    "/Documents/AI Adoption RMT/RMT_Corning/RMT_Heppner"
)

_paths: "ProjectPaths | None" = None


def is_colab() -> bool:
    try:
        import google.colab  # noqa: F401

        return True
    except ImportError:
        return False


def code_dir() -> Path:
    override = os.environ.get("HEPPNER_CODE_DIR")
    if override:
        return Path(override).expanduser().resolve()
    if COLAB_CODE_DIR.exists():
        return COLAB_CODE_DIR.resolve()
    return Path(__file__).resolve().parent


def ensure_runtime_setup() -> None:
    """Prepare sys.path and Google Drive mount when running in Colab."""
    colab_code = COLAB_CODE_DIR
    if colab_code.exists():
        code_path = str(colab_code)
        if code_path not in sys.path:
            sys.path.insert(0, code_path)

    if is_colab():
        mount_google_drive()


def mount_google_drive(force: bool = False) -> None:
    if not is_colab():
        return

    drive_root = Path("/content/drive")
    if drive_root.exists() and any(drive_root.iterdir()) and not force:
        return

    from google.colab import drive

    drive.mount("/content/drive")


def resolve_data_root() -> Path:
    override = os.environ.get("HEPPNER_DATA_ROOT")
    if override:
        return Path(override).expanduser().resolve()
    if is_colab():
        return COLAB_DATA_ROOT
    return code_dir()


@dataclass(frozen=True)
class ProjectPaths:
    environment: str
    code_dir: Path
    data_root: Path
    input_dir: Path
    processing_dir: Path
    output_dir: Path

    def ensure_directories(self) -> None:
        for directory in (self.input_dir, self.processing_dir, self.output_dir):
            directory.mkdir(parents=True, exist_ok=True)


def get_paths() -> ProjectPaths:
    global _paths
    if _paths is None:
        ensure_runtime_setup()
        root = resolve_data_root()
        environment = "colab" if is_colab() else "local"
        _paths = ProjectPaths(
            environment=environment,
            code_dir=code_dir(),
            data_root=root,
            input_dir=root / "input",
            processing_dir=root / "processing",
            output_dir=root / "output",
        )
        _paths.ensure_directories()
    return _paths


def exit_with_code(code: int) -> None:
    """Exit only on failure so Colab/Jupyter does not show SystemExit: 0 as an error."""
    if code:
        sys.exit(code)


def reset_paths() -> None:
    """Clear cached paths. Useful in tests or after changing environment variables."""
    global _paths
    _paths = None
