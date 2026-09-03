"""End-to-end pipeline: input -> processing -> matrix output.

Local:
    python run_pipeline.py

Google Colab:
    import sys
    sys.path.insert(0, "/content/Corning-Heppner")
    exec(open("/content/Corning-Heppner/run_pipeline.py").read())

Optional environment overrides (local or Colab):
    HEPPNER_DATA_ROOT - folder containing input/, processing/, output/
    HEPPNER_CODE_DIR    - folder containing the Python scripts
"""

from __future__ import annotations

import sys
from pathlib import Path

_colab_code = Path("/content/Corning-Heppner")
if _colab_code.exists() and str(_colab_code) not in sys.path:
    sys.path.insert(0, str(_colab_code))

from paths import ensure_runtime_setup, get_paths

ensure_runtime_setup()

from build_matrix import convert_file
from process_input import list_input_files, process_file, prompt_file_selection


def run_pipeline(selected_files: list[Path]) -> tuple[list[Path], list[Path]]:
    processed_paths: list[Path] = []
    output_paths: list[Path] = []

    print("\n" + "=" * 60)
    print("STEP 1/2: Clean input tabs and save to processing/")
    print("=" * 60)
    for file_path in selected_files:
        print(f"\n{file_path.name}")
        processed_path = process_file(file_path)
        processed_paths.append(processed_path)
        print(f"Saved: {processed_path}")

    print("\n" + "=" * 60)
    print("STEP 2/2: Build matrix with shipment and cost columns -> output/")
    print("=" * 60)
    for processed_path in processed_paths:
        print(f"\n{processed_path.name}")
        output_path = convert_file(processed_path)
        output_paths.append(output_path)
        print(f"Saved: {output_path}")

    return processed_paths, output_paths


def main() -> int:
    paths = get_paths()

    if not paths.input_dir.exists():
        print(f"Input folder not found: {paths.input_dir}")
        return 1

    paths.ensure_directories()

    files = list_input_files()
    if not files:
        print(f"No Excel files found in {paths.input_dir}")
        return 1

    print("Heppner tariff pipeline")
    print(f"Environment: {paths.environment}")
    print(f"Code folder: {paths.code_dir}")
    print(f"Data root:   {paths.data_root}")
    print("input/ -> processing/ -> output/")

    try:
        selected_files = prompt_file_selection(files)
    except ValueError as exc:
        print(f"Error: {exc}")
        return 1

    try:
        processed_paths, output_paths = run_pipeline(selected_files)
    except Exception as exc:
        print(f"\nPipeline failed: {exc}")
        return 1

    print("\n" + "=" * 60)
    print("PIPELINE COMPLETE")
    print("=" * 60)
    print(f"Input files processed: {len(selected_files)}")
    print(f"Intermediate files:    {len(processed_paths)}")
    print(f"Output files:          {len(output_paths)}")
    print("\nOutputs:")
    for output_path in output_paths:
        print(f"  - {output_path}")

    print("\nDone.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
