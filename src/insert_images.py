"""Handles image insertion into a python-docx Document with consistent sizing."""

import logging
from pathlib import Path

from docx.shared import Inches
from docx.oxml.ns import qn
from PIL import Image

logger = logging.getLogger(__name__)


def get_fit_dimensions(
    image_path: Path,
    max_width_inches: float,
    max_height_inches: float,
) -> tuple[float, float]:
    """
    Calculate display dimensions that fit the image within the max bounds
    while preserving the original aspect ratio.
    """
    with Image.open(image_path) as img:
        orig_w, orig_h = img.size

    # Assume 96 DPI screen resolution for inch conversion from pixels
    dpi = 96
    orig_w_in = orig_w / dpi
    orig_h_in = orig_h / dpi

    scale_w = min(1.0, max_width_inches / orig_w_in)
    scale_h = min(1.0, max_height_inches / orig_h_in)
    scale = min(scale_w, scale_h)

    return orig_w_in * scale, orig_h_in * scale


def insert_image_after_paragraph(doc_paragraph, image_path: Path, width_inches: float) -> None:
    """
    Insert an image into the document by adding a new paragraph after the given paragraph.
    The placeholder paragraph is cleared and the image is placed in it.
    """
    from docx.shared import Inches

    run = doc_paragraph.clear()
    run = doc_paragraph.add_run()
    run.add_picture(str(image_path), width=Inches(width_inches))
    logger.debug("Inserted image %s at width %.2f inches", image_path.name, width_inches)


def insert_image_into_run(run, image_path: Path, width_inches: float) -> None:
    """Replace a run's text with an inline image."""
    run.text = ""
    run.add_picture(str(image_path), width=Inches(width_inches))
    logger.debug("Inserted image %s into run at width %.2f inches", image_path.name, width_inches)
