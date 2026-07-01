"""Copies and organizes ISTA 3B input files into the output folder structure. Never deletes originals."""

import logging
import shutil
from pathlib import Path

from src.utils import sanitize_for_filename

logger = logging.getLogger(__name__)

# Sequence photo folder names that map to output sub-directories
SEQ_PHOTO_FOLDERS = ["Seq2_Tip_Over", "Seq3_Rot_Drop_1", "Seq4_Incline_1",
                     "Seq6_Rot_Drop_2", "Seq7_Incline_2"]


def build_output_structure(parsed: dict, cfg: dict) -> dict:
    """
    Build the full output folder hierarchy from parsed summary values.
    Returns a dict of named output sub-folder Paths (all created).
    """
    base = Path(cfg["output_folder"])
    customer = sanitize_for_filename(parsed.get("customer_name") or "Unknown_Customer")
    project = sanitize_for_filename(parsed.get("project_number") or "Unknown_Project")
    test_date = sanitize_for_filename(parsed.get("test_date") or "Unknown_Date")
    test_type = sanitize_for_filename(parsed.get("test_type") or "Unknown_Test")

    report_folder = base / customer / project / f"{test_date} - {test_type}"
    photos_root = report_folder / "Photos"

    paths = {
        "report_root": report_folder,
        "raw_data": report_folder / "Raw Data",
        "charts": report_folder / "Charts",
        "photos_pre": photos_root / "Pre-Test",
        "photos_post": photos_root / "Post-Test",
        "photos_accel": photos_root / "Accelerometer",
        "photos_seq2": photos_root / "Seq2_Tip_Over",
        "photos_seq3": photos_root / "Seq3_Rot_Drop_1",
        "photos_seq4": photos_root / "Seq4_Incline_1",
        "photos_seq6": photos_root / "Seq6_Rot_Drop_2",
        "photos_seq7": photos_root / "Seq7_Incline_2",
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
    Returns a manifest-ready dict of copied file name lists.
    """
    copied: dict[str, list] = {
        "raw_data": [],
        "charts": [],
        "pre_test_photos": [],
        "post_test_photos": [],
        "accel_photos": [],
        "seq_photos": {},
    }

    # ── Raw data (summary + any other non-chart files in Lansmont/) ──
    if discovered.get("summary_file"):
        _safe_copy(discovered["summary_file"], output_paths["raw_data"] / discovered["summary_file"].name)
        copied["raw_data"].append(discovered["summary_file"].name)

    if discovered.get("lansmont_folder") and discovered["lansmont_folder"].exists():
        for f in sorted(discovered["lansmont_folder"].iterdir()):
            if f.is_file() and not f.name.upper().startswith("CHART_"):
                dst = output_paths["raw_data"] / f.name
                if not dst.exists():
                    _safe_copy(f, dst)
                    if f.name not in copied["raw_data"]:
                        copied["raw_data"].append(f.name)

    # ── Charts ──
    for chart in discovered.get("charts", []):
        dst = output_paths["charts"] / chart.name
        _safe_copy(chart, dst)
        copied["charts"].append(chart.name)

    # ── Pre-test photos ──
    for photo in discovered.get("pre_test_photos", []):
        dst = output_paths["photos_pre"] / photo.name
        _safe_copy(photo, dst)
        copied["pre_test_photos"].append(photo.name)

    # ── Post-test photos ──
    for photo in discovered.get("post_test_photos", []):
        dst = output_paths["photos_post"] / photo.name
        _safe_copy(photo, dst)
        copied["post_test_photos"].append(photo.name)

    # ── Accelerometer photos ──
    for photo in discovered.get("accel_photos", []):
        dst = output_paths["photos_accel"] / photo.name
        _safe_copy(photo, dst)
        copied["accel_photos"].append(photo.name)

    # ── Per-sequence photos ──
    seq_photos: dict[str, list] = discovered.get("seq_photos", {})
    for folder_name, photos in seq_photos.items():
        output_key = _seq_folder_to_output_key(folder_name)
        seq_out = output_paths.get(output_key)
        if not seq_out:
            logger.warning("No output path for seq photo folder: %s", folder_name)
            continue
        seq_out.mkdir(parents=True, exist_ok=True)
        copied["seq_photos"][folder_name] = []
        for photo in photos:
            dst = seq_out / photo.name
            _safe_copy(photo, dst)
            copied["seq_photos"][folder_name].append(photo.name)

    logger.info(
        "Files copied — charts: %d, pre: %d, post: %d, accel: %d, seq: %s",
        len(copied["charts"]),
        len(copied["pre_test_photos"]),
        len(copied["post_test_photos"]),
        len(copied["accel_photos"]),
        {k: len(v) for k, v in copied["seq_photos"].items()},
    )
    return copied


def _seq_folder_to_output_key(folder_name: str) -> str:
    mapping = {
        "Seq2_Tip_Over": "photos_seq2",
        "Seq3_Rot_Drop_1": "photos_seq3",
        "Seq4_Incline_1": "photos_seq4",
        "Seq6_Rot_Drop_2": "photos_seq6",
        "Seq7_Incline_2": "photos_seq7",
    }
    return mapping.get(folder_name, "")


def _safe_copy(src: Path, dst: Path) -> None:
    if dst.exists():
        logger.debug("Skipping already-copied file: %s", dst.name)
        return
    shutil.copy2(src, dst)
    logger.debug("Copied %s -> %s", src.name, dst.name)
