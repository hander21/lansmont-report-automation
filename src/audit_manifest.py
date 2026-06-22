"""Creates a manifest.json audit file for every generated report."""

import json
import logging
from datetime import datetime, timezone
from pathlib import Path

logger = logging.getLogger(__name__)


def build_manifest(
    parsed: dict,
    discovered: dict,
    output_paths: dict,
    copied_files: dict,
    inserted_images: list,
    generated_report: Path,
    input_folder: Path,
    warnings: list[str],
) -> dict:
    """Assemble the full manifest dict from all inputs and outputs."""

    values_used = {}
    summary_file = discovered.get("summary_file")
    summary_name = summary_file.name if summary_file else "unknown"

    for field, value in parsed.items():
        values_used[field] = {
            "value": value,
            "source_file": summary_name,
            "source_field": field,
        }

    image_sources = {}
    for img in inserted_images:
        image_sources[img["placeholder"]] = img["file"]

    manifest = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "customer_name": parsed.get("customer_name"),
        "project_number": parsed.get("project_number"),
        "test_date": parsed.get("test_date"),
        "test_type": parsed.get("test_type"),
        "input_folder": str(input_folder.resolve()),
        "output_folder": str(output_paths["report_root"].resolve()),
        "source_files": {
            "summary_data": summary_name,
            "charts": copied_files.get("charts", []),
            "pre_test_photos": copied_files.get("pre_test_photos", []),
            "post_test_photos": copied_files.get("post_test_photos", []),
        },
        "values_used": values_used,
        "images_inserted": image_sources,
        "warnings": warnings,
        "generated_report": generated_report.name,
        "generated_report_path": str(generated_report.resolve()),
        "report_status": "DRAFT",
        "human_review_required": True,
    }
    return manifest


def write_manifest(manifest: dict, output_paths: dict) -> Path:
    """Write manifest.json into the Final Report folder."""
    manifest_path = output_paths["final_report"] / "manifest.json"
    with open(manifest_path, "w", encoding="utf-8") as f:
        json.dump(manifest, f, indent=2, default=str)
    logger.info("Manifest written: %s", manifest_path)
    return manifest_path
