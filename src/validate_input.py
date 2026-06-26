"""Validates and discovers all input files for ISTA 3A/3B report generation.

Accepts any folder structure — files are classified by filename and folder-name
keywords, not by requiring an exact directory hierarchy.
Photos are always optional: missing slots remain as placeholders in the report
so the engineer can insert them manually in Word.
"""

import logging
from pathlib import Path

from src.parse_lansmont import REQUIRED_SUMMARY_FIELDS

logger = logging.getLogger(__name__)

# Keyword → photo bucket. Checked in priority order (seq before pre/post).
_PHOTO_BUCKET_RULES = [
    ({"SEQ2", "TIP"},      "seq2"),
    ({"SEQ3"},             "seq3"),
    ({"SEQ4"},             "seq4"),
    ({"SEQ6"},             "seq6"),
    ({"SEQ7"},             "seq7"),
    ({"ACCEL"},            "accel"),
    ({"PRE"},              "pre"),
    ({"POST"},             "post"),
]

_SUMMARY_NAME_KEYWORDS  = {"SUMMARY", "TEST_SUMMARY", "LANSMONT_SUMMARY"}
_CHART_PARENT_KEYWORDS  = {"LANSMONT", "CHARTS", "CHART"}


class ValidationError(Exception):
    """Raised when required inputs are missing or invalid."""


# ─── Public API ───────────────────────────────────────────────────────────────

def validate_input_folder(input_folder: Path, cfg: dict) -> dict:
    """
    Recursively scan input_folder and classify all files.

    - Summary file:  any CSV/XLSX/TXT named SUMMARY* anywhere in the tree.
    - Charts:        any image with CHART in the filename, anywhere in the tree.
    - Photos:        classified by folder/filename keywords; ALL optional.
      Missing photo slots remain as [IMG: slot] placeholders — the engineer
      adds photos manually in Word.

    Returns a discovered dict compatible with the rest of the pipeline.
    Raises ValidationError only for missing summary or missing charts (hard errors).
    """
    if not input_folder.exists():
        raise ValidationError(f"Input folder does not exist: {input_folder.resolve()}")
    if not input_folder.is_dir():
        raise ValidationError(f"Input path is not a directory: {input_folder.resolve()}")

    file_exts = cfg["file_extensions"]
    required  = cfg["required_files"]

    summary_exts = set(e.lower() for e in file_exts.get("summary", []))
    chart_exts   = set(e.lower() for e in file_exts.get("charts",  []))
    photo_exts   = set(e.lower() for e in file_exts.get("photos",  []))

    summary_candidates: list[Path] = []
    chart_files: list[Path] = []
    photo_buckets: dict[str, list[Path]] = {
        "pre": [], "post": [], "accel": [],
        "seq2": [], "seq3": [], "seq4": [],
        "seq6": [], "seq7": [],
    }
    lansmont_folder: Path | None = None

    for f in sorted(input_folder.rglob("*")):
        if not f.is_file():
            continue
        ext      = f.suffix.lower()
        name_up  = f.name.upper()
        par_up   = f.parent.name.upper()

        # Track a folder named "Lansmont" for raw-data copy
        if "LANSMONT" in par_up and lansmont_folder is None:
            lansmont_folder = f.parent

        # ── Summary ──
        if ext in summary_exts:
            if any(kw in name_up for kw in _SUMMARY_NAME_KEYWORDS):
                summary_candidates.append(f)
                continue

        # ── Chart ──
        if ext in chart_exts and "CHART" in name_up:
            chart_files.append(f)
            continue

        # Chart in a Lansmont/Charts folder even without "CHART" in the name
        if ext in {".png", ".jpg", ".jpeg"} and any(kw in par_up for kw in _CHART_PARENT_KEYWORDS):
            if "CHART" not in name_up:  # don't double-add
                chart_files.append(f)
            continue

        # ── Photo ──
        if ext in photo_exts:
            bucket = _classify_photo(f)
            if bucket:
                photo_buckets[bucket].append(f)

    # ── Validate summary (hard error) ──
    errors:   list[str] = []
    warnings: list[str] = []

    if not summary_candidates:
        if required.get("summary_data", True):
            errors.append(
                f"No summary file found in {input_folder}. "
                "Expected a file named SUMMARY.csv (or LANSMONT_SUMMARY.csv, etc.)."
            )
        summary_file = None
    elif len(summary_candidates) > 1:
        errors.append(
            f"Multiple summary files found: {[f.name for f in summary_candidates]}. "
            "Remove duplicates."
        )
        summary_file = None
    else:
        summary_file = summary_candidates[0]
        logger.info("Summary file found: %s", summary_file.name)

    # ── Validate charts (hard error) ──
    if chart_files:
        logger.info("Charts found: %s", [f.name for f in chart_files])
    elif required.get("charts", True):
        errors.append(
            f"No chart files found in {input_folder}. "
            "Expected image files with 'CHART' in the filename."
        )

    # ── Photos — always optional ──
    for bucket, photos in photo_buckets.items():
        if photos:
            logger.info("%s photos found: %d", bucket, len(photos))
        else:
            warnings.append(
                f"No photos found for '{bucket}' slot — "
                "placeholder will remain in report for manual insertion."
            )

    for w in warnings:
        logger.warning("VALIDATION WARNING: %s", w)

    if errors:
        raise ValidationError(
            f"Input validation failed with {len(errors)} error(s):\n  - "
            + "\n  - ".join(errors)
        )

    logger.info("Input validation passed.")

    # Build seq_photos dict keyed by folder name for backwards compat
    seq_map = {
        "seq2": "Seq2_Tip_Over",
        "seq3": "Seq3_Rot_Drop_1",
        "seq4": "Seq4_Incline_1",
        "seq6": "Seq6_Rot_Drop_2",
        "seq7": "Seq7_Incline_2",
    }
    seq_photos = {
        folder_name: photo_buckets[key]
        for key, folder_name in seq_map.items()
        if photo_buckets.get(key)
    }

    return {
        "summary_file":     summary_file,
        "charts":           chart_files,
        "pre_test_photos":  photo_buckets.get("pre",   []),
        "post_test_photos": photo_buckets.get("post",  []),
        "accel_photos":     photo_buckets.get("accel", []),
        "seq_photos":       seq_photos,
        "lansmont_folder":  lansmont_folder,
        "warnings":         warnings,
    }


def validate_template(template_path: Path) -> None:
    if not template_path.exists():
        raise ValidationError(f"Report template not found: {template_path.resolve()}")
    if template_path.suffix.lower() != ".docx":
        raise ValidationError(f"Template must be a .docx file, got: {template_path.suffix}")
    logger.info("Template found: %s", template_path.name)


def validate_parsed_fields(parsed: dict) -> None:
    """Ensure all required SUMMARY fields are present and non-empty."""
    missing = [
        f for f in sorted(REQUIRED_SUMMARY_FIELDS)
        if not parsed.get(f, "").strip()
    ]
    if missing:
        raise ValidationError(
            f"Required fields missing or empty in summary data: {missing}\n"
            "Do NOT guess or estimate these values. Fix the source file."
        )
    logger.info("All required summary fields present.")


# ─── Private helpers ──────────────────────────────────────────────────────────

def _classify_photo(path: Path) -> str | None:
    """
    Classify a photo into a bucket using folder and filename keywords.
    Returns bucket name or None (photo is ignored with no error).
    """
    # Collect uppercase tokens from all path parts and the stem
    tokens: set[str] = set()
    for part in path.parts:
        for tok in part.upper().replace("-", "_").replace(" ", "_").split("_"):
            if tok:
                tokens.add(tok)

    for keywords, bucket in _PHOTO_BUCKET_RULES:
        if keywords & tokens:   # any keyword matches
            return bucket

    return None
