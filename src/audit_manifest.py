"""Creates a manifest.json audit file for every generated report."""

import json
import logging
from datetime import datetime, timezone
from pathlib import Path

logger = logging.getLogger(__name__)

# Fields sourced from SUMMARY.csv
_SUMMARY_FIELDS = {
    "customer_name", "project_number", "test_date", "test_type", "test_standard",
    "product_name", "packaging_description", "weight_lbs", "dimensions_lwh",
    "quantity", "key_takeaways", "conclusion", "technician_notes",
}


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

    summary_name = (
        discovered["summary_file"].name if discovered.get("summary_file") else "unknown"
    )

    # ── Values traceability ──
    values_used = {}

    # Summary fields
    for field in _SUMMARY_FIELDS:
        if field in parsed:
            values_used[field] = {
                "value": parsed[field],
                "source_file": summary_name,
                "source_field": field,
            }

    # Equipment fields
    equip_rows = parsed.get("_equipment_rows", [])
    for i, row in enumerate(equip_rows, start=1):
        for col, val in row.items():
            key = f"equip_{i}_{col}"
            values_used[key] = {
                "value": val,
                "source_file": "EQUIPMENT.csv",
                "source_field": col,
            }

    # Sequence fields
    for stem, seq_data in parsed.get("_sequences", {}).items():
        for field, val in seq_data.items():
            key = f"{stem.lower()}_{field}"
            values_used[key] = {
                "value": val,
                "source_file": f"{stem}.csv",
                "source_field": field,
            }

    # ── Images traceability ──
    image_sources = {img["slot"]: img["file"] for img in inserted_images}

    # ── Source files list ──
    source_files = {
        "summary_data": summary_name,
        "charts": copied_files.get("charts", []),
        "pre_test_photos": copied_files.get("pre_test_photos", []),
        "post_test_photos": copied_files.get("post_test_photos", []),
        "accel_photos": copied_files.get("accel_photos", []),
        "seq_photos": copied_files.get("seq_photos", {}),
    }

    manifest = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "customer_name": parsed.get("customer_name"),
        "project_number": parsed.get("project_number"),
        "test_date": parsed.get("test_date"),
        "test_type": parsed.get("test_type"),
        "input_folder": str(input_folder.resolve()),
        "output_folder": str(output_paths["report_root"].resolve()),
        "source_files": source_files,
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
