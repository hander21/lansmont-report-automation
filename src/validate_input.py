"""Validates that all required input files exist before report generation begins."""

import logging
from pathlib import Path

from src.utils import find_files_by_prefix

logger = logging.getLogger(__name__)

# Fields that must be present and non-empty in the parsed summary
REQUIRED_SUMMARY_FIELDS = [
    "customer_name",
    "project_number",
    "test_date",
    "test_type",
    "package_description",
    "test_standard",
    "peak_g",
    "duration",
    "axis",
    "result",
]


class ValidationError(Exception):
    """Raised when required inputs are missing or invalid."""


class ValidationWarning(Exception):
    """Raised for non-fatal issues that should be logged but not halt execution."""


def validate_input_folder(input_folder: Path, cfg: dict) -> dict:
    """
    Validate the input folder structure. Returns a dict of discovered file paths.
    Raises ValidationError on any hard failure.
    """
    errors = []
    warnings = []

    if not input_folder.exists():
        raise ValidationError(f"Input folder does not exist: {input_folder.resolve()}")
    if not input_folder.is_dir():
        raise ValidationError(f"Input path is not a directory: {input_folder.resolve()}")

    lansmont_folder = input_folder / "Lansmont"
    photos_folder = input_folder / "Photos"
    pre_test_folder = photos_folder / "Pre-Test"
    post_test_folder = photos_folder / "Post-Test"

    file_patterns = cfg["file_patterns"]
    file_exts = cfg["file_extensions"]
    required = cfg["required_files"]

    discovered = {
        "summary_file": None,
        "charts": [],
        "pre_test_photos": [],
        "post_test_photos": [],
        "lansmont_folder": lansmont_folder if lansmont_folder.exists() else None,
    }

    # --- Summary file ---
    if required.get("summary_data", True):
        if not lansmont_folder.exists():
            errors.append(f"Lansmont folder not found: {lansmont_folder}")
        else:
            summary_files = find_files_by_prefix(
                lansmont_folder,
                file_patterns["summary_prefixes"],
                file_exts["summary"],
            )
            if not summary_files:
                errors.append(
                    f"No summary file found in {lansmont_folder}. "
                    f"Expected a file starting with one of: {file_patterns['summary_prefixes']}"
                )
            elif len(summary_files) > 1:
                errors.append(
                    f"Multiple summary files found in {lansmont_folder}: "
                    f"{[f.name for f in summary_files]}. Remove duplicates or rename one."
                )
            else:
                discovered["summary_file"] = summary_files[0]
                logger.info("Summary file found: %s", summary_files[0].name)

    # --- Charts ---
    if required.get("charts", True):
        if lansmont_folder.exists():
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

    # --- Pre-test photos ---
    if required.get("pre_test_photos", True):
        if not pre_test_folder.exists():
            errors.append(f"Pre-Test photo folder not found: {pre_test_folder}")
        else:
            pre_photos = find_files_by_prefix(
                pre_test_folder,
                file_patterns["pre_test_photo_prefixes"],
                file_exts["photos"],
            )
            if not pre_photos:
                errors.append(
                    f"No pre-test photos found in {pre_test_folder}. "
                    f"Expected files starting with: {file_patterns['pre_test_photo_prefixes']}"
                )
            else:
                discovered["pre_test_photos"] = pre_photos
                logger.info("Pre-test photos found: %s", [f.name for f in pre_photos])
    else:
        if pre_test_folder.exists():
            pre_photos = find_files_by_prefix(
                pre_test_folder,
                file_patterns["pre_test_photo_prefixes"],
                file_exts["photos"],
            )
            if not pre_photos:
                warnings.append("Pre-test photos not required by config but folder is empty.")
            else:
                discovered["pre_test_photos"] = pre_photos

    # --- Post-test photos ---
    if required.get("post_test_photos", True):
        if not post_test_folder.exists():
            errors.append(f"Post-Test photo folder not found: {post_test_folder}")
        else:
            post_photos = find_files_by_prefix(
                post_test_folder,
                file_patterns["post_test_photo_prefixes"],
                file_exts["photos"],
            )
            if not post_photos:
                errors.append(
                    f"No post-test photos found in {post_test_folder}. "
                    f"Expected files starting with: {file_patterns['post_test_photo_prefixes']}"
                )
            else:
                discovered["post_test_photos"] = post_photos
                logger.info("Post-test photos found: %s", [f.name for f in post_photos])
    else:
        if post_test_folder.exists():
            post_photos = find_files_by_prefix(
                post_test_folder,
                file_patterns["post_test_photo_prefixes"],
                file_exts["photos"],
            )
            if not post_photos:
                warnings.append("Post-test photos not required by config but folder is empty.")
            else:
                discovered["post_test_photos"] = post_photos

    for w in warnings:
        logger.warning("VALIDATION WARNING: %s", w)
    discovered["warnings"] = warnings

    if errors:
        error_block = "\n  - ".join(errors)
        raise ValidationError(
            f"Input validation failed with {len(errors)} error(s):\n  - {error_block}"
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
    """Ensure all required summary fields are present and non-empty."""
    missing = []
    for field in REQUIRED_SUMMARY_FIELDS:
        value = parsed.get(field)
        if value is None or str(value).strip() == "":
            missing.append(field)

    if missing:
        raise ValidationError(
            f"Required fields missing or empty in summary data: {missing}\n"
            "Do NOT guess or estimate these values. Fix the source file."
        )
    logger.info("All required summary fields present.")
