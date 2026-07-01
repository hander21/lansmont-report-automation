"""Extract chart images from Lansmont-exported PDF files."""

import logging
from pathlib import Path

import fitz  # PyMuPDF

logger = logging.getLogger(__name__)

_VIB_NAME_KEYWORDS = {"VIB", "VIBRATION", "SEQ5"}


def is_vibration_pdf(pdf_path: Path) -> bool:
    """Return True if the PDF filename looks like a vibration test export."""
    name = pdf_path.stem.upper().replace("-", "_").replace(" ", "_")
    return any(kw in name for kw in _VIB_NAME_KEYWORDS)


def extract_charts_from_pdf(
    pdf_path: Path,
    output_dir: Path,
    vibration_only_transmissibility: bool = False,
) -> list[Path]:
    """
    Render each page of a PDF to a PNG saved in output_dir.

    vibration_only_transmissibility=True  → skip pages that don't mention
    "transmissibility" (ignores raw time-history pages, keeps the summary charts).

    Returns list of saved PNG paths, named CHART_<stem>_PAGE##.png so the
    chart-routing logic in fill_template.py can match them by keyword.
    """
    output_dir.mkdir(parents=True, exist_ok=True)
    extracted: list[Path] = []

    try:
        doc = fitz.open(str(pdf_path))
    except Exception as exc:
        logger.warning("Could not open PDF %s: %s", pdf_path.name, exc)
        return []

    stem = pdf_path.stem.upper().replace(" ", "_")

    for page_num in range(len(doc)):
        page = doc[page_num]

        if vibration_only_transmissibility:
            text = page.get_text().lower()
            if "transmissibility" not in text:
                continue

        mat = fitz.Matrix(2.0, 2.0)  # 2× scale → ~150 dpi → good quality
        pix = page.get_pixmap(matrix=mat)

        img_name = f"CHART_{stem}_PAGE{page_num + 1:02d}.png"
        img_path = output_dir / img_name
        pix.save(str(img_path))
        extracted.append(img_path)
        logger.info(
            "PDF → chart: %s page %d → %s", pdf_path.name, page_num + 1, img_name
        )

    doc.close()
    return extracted


def extract_all_pdf_charts(pdf_files: list[Path], output_dir: Path) -> list[Path]:
    """
    Process every PDF in pdf_files.
    Vibration PDFs: extract only transmissibility pages.
    All other PDFs: extract every page.

    Returns combined flat list of all extracted PNG paths.
    """
    all_charts: list[Path] = []
    for pdf_path in pdf_files:
        vib = is_vibration_pdf(pdf_path)
        charts = extract_charts_from_pdf(
            pdf_path, output_dir, vibration_only_transmissibility=vib
        )
        kind = "vibration (transmissibility only)" if vib else "drop/other (all pages)"
        logger.info(
            "PDF extracted: %s [%s] → %d chart image(s)",
            pdf_path.name,
            kind,
            len(charts),
        )
        all_charts.extend(charts)
    return all_charts
