"""Copies and organizes input files into the output folder structure. Never deletes originals."""

import logging
import shutil
from pathlib import Path

from src.utils import sanitize_for_filename

logger = logging.getLogger(__name__)


def build_output_structure(parsed: dict, cfg: dict) -> dict:
    """
    Build the full output folder path from parsed summary values.
    Returns a dict of named output sub-folder paths.
    """
    base = Path(cfg["output_folder"])
    customer = sanitize_for_filename(parsed["customer_name"])
    project = sanitize_for_filename(parsed["project_number"])
    test_date = sanitize_for_filename(parsed["test_date"])
    test_type = sanitize_for_filename(parsed["test_type"])

    report_folder = base / customer / project / f"{test_date} - {test_type}"

    paths = {
        "report_root": report_folder,
        "raw_data": report_folder / "Raw Data",
        "charts": report_folder / "Charts",
        "photos_pre": report_folder / "Photos" / "Pre-Test",
        "photos_post": report_folder / "Photos" / "Post-Test",
        "final_report": report_folder / "Final Report",
    }

    for name, path in paths.items():
        path.mkdir(parents=True, exist_ok=True)
        logger.debug("Output folder ready: %s -> %s", name, path)

    return paths


def copy_files(discovered: dict, output_paths: dict) -> dict:
    """
    Copy source files into the organized output structure.
    Original files are NEVER moved or deleted.
    Returns a manifest-ready dict of copied file paths.
    """
    copied = {
        "raw_data": [],
        "charts": [],
        "pre_test_photos": [],
        "post_test_photos": [],
    }

    # Raw data / summary
    if discovered.get("summary_file"):
        src = discovered["summary_file"]
        dst = output_paths["raw_data"] / src.name
        _safe_copy(src, dst)
        copied["raw_data"].append(dst.name)

    # Lansmont folder — copy everything from it into Raw Data (charts handled separately)
    if discovered.get("lansmont_folder") and discovered["lansmont_folder"].exists():
        for f in discovered["lansmont_folder"].iterdir():
            if f.is_file() and not f.name.lower().startswith("chart_"):
                dst = output_paths["raw_data"] / f.name
                _safe_copy(f, dst)

    # Charts
    for chart in discovered.get("charts", []):
        dst = output_paths["charts"] / chart.name
        _safe_copy(chart, dst)
        copied["charts"].append(chart.name)

    # Pre-test photos
    for photo in discovered.get("pre_test_photos", []):
        dst = output_paths["photos_pre"] / photo.name
        _safe_copy(photo, dst)
        copied["pre_test_photos"].append(photo.name)

    # Post-test photos
    for photo in discovered.get("post_test_photos", []):
        dst = output_paths["photos_post"] / photo.name
        _safe_copy(photo, dst)
        copied["post_test_photos"].append(photo.name)

    logger.info(
        "Files copied — charts: %d, pre-photos: %d, post-photos: %d",
        len(copied["charts"]),
        len(copied["pre_test_photos"]),
        len(copied["post_test_photos"]),
    )
    return copied


def _safe_copy(src: Path, dst: Path) -> None:
    if dst.exists():
        logger.debug("Skipping already-copied file: %s", dst.name)
        return
    shutil.copy2(src, dst)
    logger.debug("Copied %s -> %s", src, dst)
