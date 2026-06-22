"""Generates a draft Word report from the template with values and images inserted."""

import logging
from pathlib import Path

from docx import Document
from docx.shared import Inches, Pt, RGBColor

from src.insert_images import get_fit_dimensions
from src.utils import sanitize_for_filename

logger = logging.getLogger(__name__)

# Placeholder tokens used in the template
PLACEHOLDER_FIELDS = [
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
    "technician_notes",
    "conclusion",
    "draft_watermark",
]

IMAGE_PLACEHOLDERS = {
    "{{ vibration_chart }}": "vibration_chart",
    "{{ incline_impact_chart }}": "incline_impact_chart",
    "{{ pre_test_photo_1 }}": "pre_test_photo_1",
    "{{ pre_test_photo_2 }}": "pre_test_photo_2",
    "{{ post_test_photo_1 }}": "post_test_photo_1",
    "{{ post_test_photo_2 }}": "post_test_photo_2",
}


def build_template_context(parsed: dict, cfg: dict) -> dict:
    """Assemble the text context dict from parsed summary values."""
    ctx = {}
    for field in PLACEHOLDER_FIELDS:
        if field == "draft_watermark":
            ctx[field] = "** DRAFT — NOT FOR DISTRIBUTION **"
        elif field == "conclusion":
            ctx[field] = parsed.get(
                "conclusion",
                "DRAFT CONCLUSION: Review all sections before finalizing.",
            )
        elif field == "technician_notes":
            ctx[field] = parsed.get("technician_notes", "(No technician notes provided)")
        else:
            ctx[field] = parsed.get(field, "")
    return ctx


def _replace_text_in_paragraph(paragraph, placeholder: str, replacement: str) -> None:
    """Replace placeholder text in a paragraph while preserving run formatting."""
    full_text = "".join(r.text for r in paragraph.runs)
    if placeholder not in full_text:
        return

    new_text = full_text.replace(placeholder, replacement)

    # Clear all runs and put the text into the first one
    for i, run in enumerate(paragraph.runs):
        run.text = new_text if i == 0 else ""


def _assign_charts(discovered: dict, output_paths: dict) -> dict:
    """Map logical chart names to their output-folder file paths."""
    charts_out = {}
    chart_files = [
        output_paths["charts"] / f
        for f in output_paths["charts"].iterdir()
        if output_paths["charts"].is_dir()
    ] if output_paths["charts"].exists() else []

    # Re-read from the charts output folder
    chart_paths = sorted(output_paths["charts"].glob("*")) if output_paths["charts"].exists() else []

    for cp in chart_paths:
        name_lower = cp.name.lower()
        if "vibration" in name_lower:
            charts_out.setdefault("vibration_chart", cp)
        elif "incline" in name_lower or "impact" in name_lower:
            charts_out.setdefault("incline_impact_chart", cp)
        else:
            logger.warning("Chart file not mapped to a placeholder: %s", cp.name)

    return charts_out


def _assign_photos(output_paths: dict) -> dict:
    """Map logical photo slot names to file paths from output folders."""
    photos = {}

    pre_photos = sorted(output_paths["photos_pre"].glob("*")) if output_paths["photos_pre"].exists() else []
    for i, p in enumerate(pre_photos[:2], start=1):
        photos[f"pre_test_photo_{i}"] = p

    post_photos = sorted(output_paths["photos_post"].glob("*")) if output_paths["photos_post"].exists() else []
    for i, p in enumerate(post_photos[:2], start=1):
        photos[f"post_test_photo_{i}"] = p

    return photos


def generate_draft_report(
    parsed: dict,
    discovered: dict,
    output_paths: dict,
    cfg: dict,
) -> Path:
    """
    Build the draft DOCX report from the template.
    Returns the path to the generated file.
    """
    template_path = Path(cfg["template_path"])
    doc = Document(str(template_path))

    ctx = build_template_context(parsed, cfg)
    charts = _assign_charts(discovered, output_paths)
    photos = _assign_photos(output_paths)
    image_map = {**charts, **photos}

    chart_cfg = cfg["chart_settings"]
    photo_cfg = cfg["photo_settings"]

    inserted_images = []
    replaced_text_placeholders = set()

    for para in doc.paragraphs:
        full_text = "".join(r.text for r in para.runs)

        # Try image placeholders first
        matched_image = False
        for token, slot_name in IMAGE_PLACEHOLDERS.items():
            if token in full_text:
                image_path = image_map.get(slot_name)
                if image_path and Path(image_path).exists():
                    is_chart = "chart" in slot_name
                    max_w = chart_cfg["max_width_inches"] if is_chart else photo_cfg["max_width_inches"]
                    max_h = chart_cfg["max_height_inches"] if is_chart else photo_cfg["max_height_inches"]
                    w, _ = get_fit_dimensions(image_path, max_w, max_h)
                    # Clear the paragraph and insert image
                    for run in para.runs:
                        run.text = ""
                    para.runs[0].add_picture(str(image_path), width=Inches(w))
                    logger.info("Inserted image for placeholder '%s': %s", token, image_path.name)
                    inserted_images.append({"placeholder": token, "file": image_path.name})
                else:
                    logger.warning(
                        "Image placeholder '%s' found but no image available for slot '%s'.",
                        token,
                        slot_name,
                    )
                matched_image = True
                break

        if matched_image:
            continue

        # Text field replacements
        for field, value in ctx.items():
            token = f"{{{{ {field} }}}}"
            if token in full_text:
                _replace_text_in_paragraph(para, token, str(value))
                replaced_text_placeholders.add(field)

    # Build output filename
    customer = sanitize_for_filename(parsed["customer_name"])
    project = sanitize_for_filename(parsed["project_number"])
    date = sanitize_for_filename(parsed["test_date"])
    test_type = sanitize_for_filename(parsed["test_type"])
    filename = f"DRAFT_{customer}_{project}_{test_type}_{date}.docx"
    output_file = output_paths["final_report"] / filename

    doc.save(str(output_file))
    logger.info("Draft report saved: %s", output_file)

    return output_file, inserted_images
