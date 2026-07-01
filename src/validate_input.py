"""Validate and discover all input files for ISTA report generation.

Accepts any folder structure — files are classified by filename keywords,
not by requiring an exact directory hierarchy.

PDFs are treated as chart sources: images are extracted from them
automatically.  Vibration PDFs yield only transmissibility pages; all
other PDFs yield every page.

Summary file, charts, and photos are all optional — missing slots stay as
placeholders in the report for the engineer to fill in manually.
"""

import logging
import tempfile
from pathlib import Path

from src.parse_lansmont import REQUIRED_SUMMARY_FIELDS

logger = logging.getLogger(__name__)

# Keyword → photo bucket.  Checked in priority order (seq before pre/post).
_PHOTO_BUCKET_RULES = [
    ({"SEQ2", "TIP"},  "seq2"),
    ({"SEQ3"},         "seq3"),
    ({"SEQ4"},         "seq4"),
    ({"SEQ6"},         "seq6"),
    ({"SEQ7"},         "seq7"),
    ({"ACCEL"},        "accel"),
    ({"PRE"},          "pre"),
    ({"POST"},         "post"),
]

_SUMMARY_NAME_KEYWORDS = {"SUMMARY", "TEST_SUMMARY", "LANSMONT_SUMMARY"}


class ValidationError(Exception):
    """Raised when inputs are fatally broken (e.g. ZIP can't be read)."""


# ─── Public API ───────────────────────────────────────────────────────────────

def validate_input_folder(input_folder: Path, cfg: dict) -> dict:
    """
    Recursively scan input_folder and classify all files.

    - Summary:  any CSV/XLSX/TXT named SUMMARY* anywhere in the tree
                (optional — missing produces a warning, not an error).
    - PDFs:     rendered to chart PNGs automatically.
    - Charts:   any image with CHART in the name, anywhere in the tree.
    - Photos:   classified by folder/filename keywords; always optional.

    Raises ValidationError only for unreadable input folders.
    All other missing files are collected as warnings and returned in the dict.
    """
    if not input_folder.exists():
        raise ValidationError(f"Input folder does not exist: {input_folder.resolve()}")
    if not input_folder.is_dir():
        raise ValidationError(f"Input path is not a directory: {input_folder.resolve()}")

    file_exts = cfg["file_extensions"]
    summary_exts = {e.lower() for e in file_exts.get("summary", [])}
    chart_exts   = {e.lower() for e in file_exts.get("charts",  [])}
    photo_exts   = {e.lower() for e in file_exts.get("photos",  [])}

    summary_candidates: list[Path] = []
    chart_files: list[Path] = []
    pdf_files: list[Path] = []
    photo_buckets: dict[str, list[Path]] = {
        "pre": [], "post": [], "accel": [],
        "seq2": [], "seq3": [], "seq4": [],
        "seq6": [], "seq7": [],
    }
    lansmont_folder: Path | None = None

    for f in sorted(input_folder.rglob("*")):
        if not f.is_file():
            continue
        ext     = f.suffix.lower()
        name_up = f.name.upper()
        par_up  = f.parent.name.upper()

        if "LANSMONT" in par_up and lansmont_folder is None:
            lansmont_folder = f.parent

        # ── Summary ──
        if ext in summary_exts and any(kw in name_up for kw in _SUMMARY_NAME_KEYWORDS):
            summary_candidates.append(f)
            continue

        # ── PDF → extract charts later ──
        if ext == ".pdf":
            pdf_files.append(f)
            continue

        # ── Chart image ──
        if ext in chart_exts and "CHART" in name_up:
            chart_files.append(f)
            continue

        # Image in a folder whose name suggests it holds charts
        if ext in {".png", ".jpg", ".jpeg"} and any(
            kw in par_up for kw in {"LANSMONT", "CHARTS", "CHART"}
        ):
            if "CHART" not in name_up:
                chart_files.append(f)
            continue

        # ── Photo ──
        if ext in photo_exts:
            bucket = _classify_photo(f)
            if bucket:
                photo_buckets[bucket].append(f)

    warnings: list[str] = []

    # ── Extract charts from PDFs ──
    if pdf_files:
        logger.info("Found %d PDF(s) — extracting chart images…", len(pdf_files))
        pdf_chart_dir = Path(tempfile.mkdtemp(prefix="lansmont_pdf_charts_"))
        try:
            from src.extract_pdf_charts import extract_all_pdf_charts
            extracted = extract_all_pdf_charts(pdf_files, pdf_chart_dir)
            chart_files.extend(extracted)
            logger.info("Extracted %d chart image(s) from PDF(s).", len(extracted))
        except Exception as exc:
            warnings.append(f"PDF chart extraction failed: {exc}")

    # ── Summary ──
    if not summary_candidates:
        warnings.append(
            "No summary file (SUMMARY.csv / .xlsx / .txt) found. "
            "Report fields will be blank — fill them in manually in Word."
        )
        summary_file = None
    elif len(summary_candidates) > 1:
        warnings.append(
            f"Multiple summary files found: {[f.name for f in summary_candidates]}. "
            f"Using the first one: {summary_candidates[0].name}"
        )
        summary_file = summary_candidates[0]
    else:
        summary_file = summary_candidates[0]
        logger.info("Summary file found: %s", summary_file.name)

    # ── Charts ──
    if chart_files:
        logger.info("Charts found: %s", [f.name for f in chart_files])
    else:
        warnings.append(
            "No chart files found. "
            "Chart slots will remain as placeholders for manual insertion."
        )

    # ── Photos ──
    for bucket, photos in photo_buckets.items():
        if photos:
            logger.info("%s photos found: %d", bucket, len(photos))
        else:
            warnings.append(
                f"No photos for '{bucket}' slot — "
                "placeholder will remain for manual insertion."
            )

    for w in warnings:
        logger.warning("VALIDATION WARNING: %s", w)

    logger.info("Input scan complete.")

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
    """Warn (but do not raise) when required summary fields are missing."""
    missing = [
        f for f in sorted(REQUIRED_SUMMARY_FIELDS)
        if not parsed.get(f, "").strip()
    ]
    if missing:
        logger.warning(
            "Some required fields are blank (fill in manually in Word): %s", missing
        )
    else:
        logger.info("All required summary fields present.")


# ─── Private helpers ──────────────────────────────────────────────────────────

def _classify_photo(path: Path) -> str | None:
    tokens: set[str] = set()
    for part in path.parts:
        for tok in part.upper().replace("-", "_").replace(" ", "_").split("_"):
            if tok:
                tokens.add(tok)
    for keywords, bucket in _PHOTO_BUCKET_RULES:
        if keywords & tokens:
            return bucket
    return None
