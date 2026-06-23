"""Validates that all required ISTA 3B input files exist before report generation begins."""

import logging
from pathlib import Path

from src.utils import find_files_by_prefix
from src.parse_lansmont import REQUIRED_SUMMARY_FIELDS

logger = logging.getLogger(__name__)

# Photo subfolders expected for ISTA 3B (sequence folders are optional — produce warnings)
REQUIRED_PHOTO_FOLDERS = ["Pre-Test", "Post-Test"]
OPTIONAL_PHOTO_FOLDERS = [
    "Accelerometer",
    "Seq2_Tip_Over",
    "Seq3_Rot_Drop_1",
    "Seq4_Incline_1",
    "Seq6_Rot_Drop_2",
    "Seq7_Incline_2",
]


class ValidationError(Exception):
    """Raised when required inputs are missing or invalid."""


def validate_input_folder(input_folder: Path, cfg: dict) -> dict:
    """
    Validate the ISTA 3B input folder structure.
    Returns a dict of discovered file paths and lists.
    Raises ValidationError on any hard failure; collects warnings for soft issues.
    """
    errors = []
    warnings = []

    if not input_folder.exists():
        raise ValidationError(f"Input folder does not exist: {input_folder.resolve()}")
    if not input_folder.is_dir():
        raise ValidationError(f"Input path is not a directory: {input_folder.resolve()}")

    lansmont_folder = input_folder / "Lansmont"
    photos_folder = input_folder / "Photos"

    file_patterns = cfg["file_patterns"]
    file_exts = cfg["file_extensions"]
    required = cfg["required_files"]

    discovered = {
        "summary_file": None,
        "charts": [],
        "pre_test_photos": [],
        "post_test_photos": [],
        "accel_photos": [],
        "seq_photos": {},        # stem → [Path, ...]
        "lansmont_folder": lansmont_folder if lansmont_folder.exists() else None,
    }

    # ── Lansmont folder ──
    if not lansmont_folder.exists():
        errors.append(f"Lansmont folder not found: {lansmont_folder}")
    else:
        # Summary file
        if required.get("summary_data", True):
            summary_files = find_files_by_prefix(
                lansmont_folder,
                file_patterns["summary_prefixes"],
                file_exts["summary"],
            )
            if not summary_files:
                errors.append(
                    f"No summary file found in {lansmont_folder}. "
                    f"Expected a file starting with: {file_patterns['summary_prefixes']}"
                )
            elif len(summary_files) > 1:
                errors.append(
                    f"Multiple summary files found in {lansmont_folder}: "
                    f"{[f.name for f in summary_files]}. Remove duplicates."
                )
            else:
                discovered["summary_file"] = summary_files[0]
                logger.info("Summary file found: %s", summary_files[0].name)

        # Charts
        if required.get("charts", True):
            charts = find_files_by_prefix(
                lansmont_folder,
                file_patterns["chart_prefixes"],
                file_exts["charts"],
            )
            if not charts:
                errors.append(
                    f"No chart files found in {lansmont_folder}. "
                    f"Expected files starting with: {file_patterns['chart_prefixes']}"
                )
            else:
                discovered["charts"] = charts
                logger.info("Charts found: %s", [f.name for f in charts])

    # ── Photos folder ──
    if not photos_folder.exists():
        errors.append(f"Photos folder not found: {photos_folder}")
    else:
        # Required photo folders
        for folder_name in REQUIRED_PHOTO_FOLDERS:
            folder = photos_folder / folder_name
            prefix_key = "pre_test_photo_prefixes" if "Pre" in folder_name else "post_test_photo_prefixes"
            discovered_key = "pre_test_photos" if "Pre" in folder_name else "post_test_photos"

            if required.get(folder_name.lower().replace("-", "_"), True):
                if not folder.exists():
                    errors.append(f"{folder_name} photo folder not found: {folder}")
                else:
                    photos = find_files_by_prefix(
                        folder, file_patterns[prefix_key], file_exts["photos"]
                    )
                    if not photos:
                        errors.append(
                            f"No photos found in {folder}. "
                            f"Expected files starting with: {file_patterns[prefix_key]}"
                        )
                    else:
                        discovered[discovered_key] = photos
                        logger.info("%s photos found: %d", folder_name, len(photos))

        # Optional sequence/accel photo folders
        for folder_name in OPTIONAL_PHOTO_FOLDERS:
            folder = photos_folder / folder_name
            if not folder.exists():
                warnings.append(f"Optional photo folder not found (will skip): {folder_name}/")
                continue
            photos = [p for p in sorted(folder.iterdir())
                      if p.is_file() and p.suffix.lower() in file_exts["photos"]]
            if not photos:
                warnings.append(f"No photos found in optional folder {folder_name}/ — skipping.")
            else:
                if folder_name == "Accelerometer":
                    discovered["accel_photos"] = photos
                    logger.info("Accelerometer photos found: %d", len(photos))
                else:
                    discovered["seq_photos"][folder_name] = photos
                    logger.info("%s photos found: %d", folder_name, len(photos))

    for w in warnings:
        logger.warning("VALIDATION WARNING: %s", w)
    discovered["warnings"] = warnings

    if errors:
        raise ValidationError(
            f"Input validation failed with {len(errors)} error(s):\n  - "
            + "\n  - ".join(errors)
        )

    logger.info("Input validation passed.")
    return discovered


def validate_template(template_path: Path) -> None:
    if not template_path.exists():
        raise ValidationError(f"Report template not found: {template_path.resolve()}")
    if template_path.suffix.lower() != ".docx":
        raise ValidationError(
            f"Template must be a .docx file, got: {template_path.suffix}"
        )
    logger.info("Template found: %s", template_path.name)


def validate_parsed_fields(parsed: dict) -> None:
    """Ensure all required SUMMARY fields are present and non-empty in the parsed dict."""
    missing = []
    for field in sorted(REQUIRED_SUMMARY_FIELDS):
        value = parsed.get(field)
        if value is None or str(value).strip() == "":
            missing.append(field)

    if missing:
        raise ValidationError(
            f"Required fields missing or empty in summary data: {missing}\n"
            "Do NOT guess or estimate these values. Fix the source file."
        )
    logger.info("All required summary fields present.")
