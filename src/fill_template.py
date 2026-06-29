"""
Fill official TransPak Word templates with test data.

Supports ISTA 2A, 3A, 3B, and Simplified templates (the real .docx files,
not the fake token-based templates used during development).

Strategy:
  1. Replace known placeholder paragraphs (product name, operators, date, revision)
  2. Refill the Equipment table from EQUIPMENT.csv data
  3. Update the Test Sample Description table from SUMMARY.csv
  4. Insert pre-test, accelerometer, and post-test photos
  5. Insert Lansmont transmissibility charts
"""

import logging
from pathlib import Path

from docx import Document
from docx.shared import Inches, Pt
from docx.oxml.ns import qn

from src.insert_images import get_fit_dimensions

logger = logging.getLogger(__name__)


# ─── Known paragraph placeholder strings in the real templates ────────────────

_PARA_REPLACEMENTS = {
    # 3B
    "Product Name":            "product_name",
    "ISTA 3B Test Method":     "test_method",
    "Rev # (A, B, C, etc)":   "revision",
    # 3A
    "Product":                 "product_name",
    "Modified ISTA 3A Test Protocol": "test_method",
    # Simplified
    "Rev (A, B, C, etc)":     "revision",
    # 2A (uses example data as placeholder)
    "Camera Pack":             "product_name",
    # Operator lines (all templates — match both placeholder and 2A example text)
    "Test Operator #1 Name | Email":            "operator_1",
    " \tTest Operator #1 Name | Email ":        "operator_1",
    "Test Operator #2 Name | Email":            "operator_2",
    "\tTest Operator #2 Name | Email":          "operator_2",
    "Test Operator #3 Name | Email":            "operator_3",
    "Jack Parent | jack.parent@transpak.com":   "operator_1",
    "Annie Schaubel | annie.schaubel@transpak.com": "operator_2",
    # Revision (3A/2A already have "Rev A" as placeholder)
    "Rev A":                   "revision",
}

_OBJECTIVE_STARTS = (
    "To validate the ability of the packaging",
    "To evaluate",
    "This test was performed",
)

# Equipment column header → CSV field name
_EQUIP_COL_MAP = {
    "equipment":            "equipment",
    "make":                 "make",
    "model":                "model",
    "s/n":                  "serial",
    "serial":               "serial",
    "calibration date":     "calibration_date",
    "calibration due date": "calibration_date",
}


# ─── Public API ───────────────────────────────────────────────────────────────

def is_real_template(doc: Document) -> bool:
    """Return True when the template does NOT use {{ }} tokens (i.e. it's a real TransPak file)."""
    for para in doc.paragraphs:
        if "{{" in para.text:
            return False
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    if "{{" in para.text:
                        return False
    return True


def fill_real_template(
    doc: Document,
    parsed: dict,
    discovered: dict,
    cfg: dict,
) -> list[dict]:
    """
    Fill a real TransPak DOCX template with parsed test data.
    Returns a list of inserted-image info dicts (same format as generate_report.py).
    """
    inserted: list[dict] = []

    _fill_header_paragraphs(doc, parsed)
    _fill_equipment_table(doc, parsed.get("_equipment_rows", []))
    _fill_sample_description(doc, parsed)

    photo_cfg  = cfg["photo_settings"]
    chart_cfg  = cfg["chart_settings"]
    max_pw     = photo_cfg["max_width_inches"]
    max_ph     = photo_cfg["max_height_inches"]
    max_cw     = chart_cfg["max_width_inches"]
    max_ch     = chart_cfg["max_height_inches"]

    accel_photos     = discovered.get("accel_photos", [])
    pre_photos       = discovered.get("pre_test_photos", [])
    post_photos      = discovered.get("post_test_photos", [])
    seq_photos       = discovered.get("seq_photos", {})
    charts           = discovered.get("charts", [])

    # ── Accelerometer photos (3B only — tables starting with "Accelerometer #N") ──
    for n, keyword in [(0, "Accelerometer #1"), (1, "Accelerometer #2")]:
        table = _find_table_by_first_cell(doc, keyword)
        if table and n < len(accel_photos):
            result = _insert_image_into_table_row(table, 1, 0, accel_photos[n], max_pw, max_ph)
            if result:
                result["kind"] = "IMG"
                result["slot"] = f"accel_{n+1}"
                inserted.append(result)

    # ── Pre-test photos (table with "Pretest Photos" header OR 3x3 empty table) ──
    pre_table = _find_table_by_first_cell(doc, "Pretest Photos") \
             or _find_table_after_heading(doc, "Pre-test Photos")
    if pre_table and pre_photos:
        inserted += _fill_photo_table(pre_table, pre_photos, max_pw, max_ph, "pre", skip_rows=1)

    # ── Sequence-specific photos (insert into matching tables by seq key) ──
    seq_table_keywords = {
        "Seq2_Tip_Over":      ["Tip/Tip Over"],
        "Seq3_Rot_Drop_1":    ["Rotational Drop"],
        "Seq4_Incline_1":     ["Incline Impact"],
        "Seq6_Rot_Drop_2":    ["Rotational Drop"],
        "Seq7_Incline_2":     ["Incline Impact"],
    }
    # We fill seq photos into unlabelled photo tables that appear after the sequence heading
    seq_heading_map = {
        "Seq2_Tip_Over":   "Test Sequence 2",
        "Seq3_Rot_Drop_1": "Test Sequence 3",
        "Seq4_Incline_1":  "Test Sequence 4",
        "Seq6_Rot_Drop_2": "Test Sequence 6",
        "Seq7_Incline_2":  "Test Sequence 7",
    }
    for seq_key, photos_list in seq_photos.items():
        if not photos_list:
            continue
        heading = seq_heading_map.get(seq_key)
        if heading:
            table = _find_empty_table_after_heading(doc, heading)
            if table:
                inserted += _fill_photo_table(table, photos_list, max_pw, max_ph, seq_key, skip_rows=0)

    # ── Post-test photos ──
    post_table = _find_table_after_heading(doc, "Post-Test Photos")
    if post_table and post_photos:
        inserted += _fill_photo_table(post_table, post_photos, max_pw, max_ph, "post", skip_rows=0)

    # ── Transmissibility / vibration charts ──
    inserted += _fill_transmissibility_charts(doc, charts, max_cw, max_ch)

    return inserted


# ─── Header paragraph replacement ────────────────────────────────────────────

def _fill_header_paragraphs(doc: Document, parsed: dict):
    """Replace known placeholder paragraphs in the cover/header area."""
    for para in doc.paragraphs:
        raw  = para.text          # includes tabs/whitespace
        text = raw.strip()

        # Test Date line (has a Word form field after "Test Date:")
        if "Test Date:" in text:
            date_val = parsed.get("test_date", "")
            if date_val:
                # Append date as a new run; this works even if a date field exists
                if date_val not in raw:
                    run = para.add_run(f"   {date_val}")
                    if para.runs:
                        try:
                            run.font.size = para.runs[0].font.size
                        except Exception:
                            pass
            continue

        # Objective paragraph (long text, only update if user provided one)
        if any(text.startswith(pfx) for pfx in _OBJECTIVE_STARTS):
            objective = parsed.get("objective", "")
            if objective:
                _set_para_text(para, objective)
            continue

        # Direct placeholder / example-text matches
        if text in _PARA_REPLACEMENTS or raw in _PARA_REPLACEMENTS:
            field = _PARA_REPLACEMENTS.get(text) or _PARA_REPLACEMENTS.get(raw)
            value = parsed.get(field, "")
            if value:
                _set_para_text(para, value)


def _set_para_text(para, new_text: str):
    """Replace paragraph text, preserving the formatting of the first run."""
    if para.runs:
        para.runs[0].text = new_text
        for run in para.runs[1:]:
            run.text = ""
    else:
        para.add_run(new_text)


# ─── Equipment table ──────────────────────────────────────────────────────────

def _fill_equipment_table(doc: Document, equipment_rows: list[dict]):
    """Clear and refill the Equipment table from parsed EQUIPMENT.csv rows."""
    if not equipment_rows:
        return

    table = _find_table_by_first_cell(doc, "Equipment Used")
    if not table:
        logger.warning("Equipment table not found in template.")
        return

    # Row 0 = merged "Equipment Used" header; Row 1 = column headers
    if len(table.rows) < 2:
        return

    # Determine column order from row 1 column headers
    col_headers = [c.text.strip().lower() for c in table.rows[1].cells]

    # Remove all data rows (keep rows 0 and 1)
    for row in list(table.rows)[2:]:
        _remove_row(row)

    # Add new rows from CSV
    for eq in equipment_rows:
        values = []
        for col in col_headers:
            csv_field = _EQUIP_COL_MAP.get(col, col)
            values.append(eq.get(csv_field, ""))
        _append_table_row(table, values, font_size=12)

    logger.info("Equipment table filled: %d row(s).", len(equipment_rows))


# ─── Test Sample Description table ───────────────────────────────────────────

def _fill_sample_description(doc: Document, parsed: dict):
    """Update the Test Sample Description data row."""
    # Find by "External Packaging" first cell, or "Packaged-Product Materials" (2A)
    table = (_find_table_by_first_cell(doc, "External Packaging")
             or _find_table_by_first_cell(doc, "Packaged-Product Materials"))
    if not table or len(table.rows) < 2:
        return

    # Row 0 = column headers; Row 1 = data
    data_row = table.rows[1]
    col_headers = [c.text.strip().lower() for c in table.rows[0].cells]

    field_map = {
        "external packaging":          "packaging_description",
        "packaged-product materials":  "packaging_description",
        "weight":                      "weight_lbs",
        "total weight":                "weight_lbs",
        "outside dimensions\n(l x w x h)": "dimensions_lwh",
        "outside dimensions":          "dimensions_lwh",
        "quantity":                    "quantity",
    }

    seen_tcs = set()
    for i, (cell, col_h) in enumerate(zip(data_row.cells, col_headers)):
        tc_id = id(cell._tc)
        if tc_id in seen_tcs:
            continue
        seen_tcs.add(tc_id)
        csv_field = field_map.get(col_h)
        if csv_field:
            value = parsed.get(csv_field, "")
            if value:
                _set_cell_text(cell, value)

    logger.info("Test sample description updated.")


# ─── Photo table filling ──────────────────────────────────────────────────────

def _fill_photo_table(
    table,
    photos: list[Path],
    max_w: float,
    max_h: float,
    slot_prefix: str,
    skip_rows: int = 1,
) -> list[dict]:
    """Insert photos into cell slots in a table, skipping header rows."""
    inserted = []
    photo_iter = iter(photos)
    seen_tcs: set[int] = set()

    for row_idx, row in enumerate(table.rows):
        if row_idx < skip_rows:
            continue
        # Skip rows that are clearly caption / text rows (all cells have significant text)
        cell_texts = [c.text.strip() for c in row.cells]
        if all(len(t) > 15 for t in cell_texts if t):
            continue

        for cell in row.cells:
            tc_id = id(cell._tc)
            if tc_id in seen_tcs:
                continue
            seen_tcs.add(tc_id)

            photo = next(photo_iter, None)
            if photo is None:
                return inserted

            result = _insert_image_in_cell(cell, photo, max_w, max_h)
            if result:
                result["kind"] = "IMG"
                result["slot"] = f"{slot_prefix}_{len(inserted)+1}"
                inserted.append(result)

    return inserted


def _insert_image_into_table_row(
    table,
    row_idx: int,
    col_idx: int,
    image_path: Path,
    max_w: float,
    max_h: float,
) -> dict | None:
    """Insert an image into a specific table cell (row_idx, col_idx)."""
    if row_idx >= len(table.rows):
        return None
    row = table.rows[row_idx]
    seen_tcs: set[int] = set()
    real_col = 0
    for cell in row.cells:
        tc_id = id(cell._tc)
        if tc_id in seen_tcs:
            continue
        seen_tcs.add(tc_id)
        if real_col == col_idx:
            return _insert_image_in_cell(cell, image_path, max_w, max_h)
        real_col += 1
    return None


# ─── Transmissibility chart insertion ────────────────────────────────────────

def _fill_transmissibility_charts(
    doc: Document,
    charts: list[Path],
    max_w: float,
    max_h: float,
) -> list[dict]:
    """Insert chart files into all tables whose header contains 'Transmissibility'."""
    inserted = []
    chart_iter = iter(charts)

    for table in doc.tables:
        if not table.rows:
            continue
        header_text = table.rows[0].cells[0].text.strip()
        if "Transmissibility" not in header_text and "Transmissibility" not in header_text:
            continue
        if len(table.rows) < 2:
            continue

        chart = next(chart_iter, None)
        if chart is None:
            break

        result = _insert_image_into_table_row(table, 1, 0, chart, max_w, max_h)
        if result:
            result["kind"] = "CHART"
            result["slot"] = header_text[:40]
            inserted.append(result)
            logger.info("Chart inserted: %s → %s", chart.name, header_text[:40])

    return inserted


# ─── Table / document body helpers ───────────────────────────────────────────

def _find_table_by_first_cell(doc: Document, keyword: str):
    """Return the first table whose first cell text starts with keyword."""
    for table in doc.tables:
        if not table.rows:
            continue
        first = table.rows[0].cells[0].text.strip()
        if first == keyword or first.startswith(keyword):
            return table
    return None


def _get_body_elements(doc: Document) -> list[tuple[str, object]]:
    """Return doc body elements in order as (kind, obj) where kind is 'para' or 'table'."""
    para_map  = {p._p: p for p in doc.paragraphs}
    table_map = {t._tbl: t for t in doc.tables}
    result = []
    for child in doc.element.body:
        local = child.tag.split("}")[-1] if "}" in child.tag else child.tag
        if local == "p" and child in para_map:
            result.append(("para", para_map[child]))
        elif local == "tbl" and child in table_map:
            result.append(("table", table_map[child]))
    return result


def _find_table_after_heading(doc: Document, heading_text: str):
    """Return the first table that appears after a paragraph with heading_text."""
    found_heading = False
    for kind, elem in _get_body_elements(doc):
        if kind == "para" and elem.text.strip() == heading_text:
            found_heading = True
        elif kind == "table" and found_heading:
            return elem
    return None


def _find_empty_table_after_heading(doc: Document, heading_text: str):
    """Return the first all-empty table after a paragraph matching heading_text."""
    found_heading = False
    for kind, elem in _get_body_elements(doc):
        if kind == "para" and heading_text in elem.text:
            found_heading = True
        elif kind == "table" and found_heading:
            # Check if table is "empty" (no meaningful text — cells may have images)
            all_cell_texts = [
                c.text.strip()
                for row in elem.rows
                for c in row.cells
            ]
            if not any(len(t) > 20 for t in all_cell_texts):
                return elem
    return None


# ─── Cell helpers ─────────────────────────────────────────────────────────────

def _insert_image_in_cell(cell, image_path: Path, max_w: float, max_h: float) -> dict | None:
    """Clear cell and insert an image. Returns info dict or None on failure."""
    if not image_path.exists():
        logger.warning("Image file missing: %s", image_path)
        return None
    try:
        _clear_cell(cell)
        para = cell.paragraphs[0] if cell.paragraphs else cell.add_paragraph()
        run  = para.add_run()
        w, _ = get_fit_dimensions(image_path, max_w, max_h)
        run.add_picture(str(image_path), width=Inches(w))
        return {"file": image_path.name}
    except Exception as e:
        logger.warning("Failed to insert image %s: %s", image_path.name, e)
        return None


def _clear_cell(cell):
    """Remove all content (text and inline images) from a cell."""
    tc = cell._tc
    # Remove all runs (and their embedded pictures) from every paragraph
    for p_elem in tc.findall(qn("w:p")):
        for r_elem in list(p_elem.findall(qn("w:r"))):
            p_elem.remove(r_elem)
        # Also remove any direct drawing elements not in a run
        for drawing in list(p_elem.findall(".//" + qn("w:drawing"))):
            drawing.getparent().remove(drawing)

    # Remove all paragraphs after the first one
    all_paras = tc.findall(qn("w:p"))
    for p_elem in all_paras[1:]:
        tc.remove(p_elem)


def _set_cell_text(cell, text: str):
    """Set the text of the first paragraph in a cell, preserving paragraph properties."""
    for i, para in enumerate(cell.paragraphs):
        if i == 0:
            if para.runs:
                para.runs[0].text = text
                for run in para.runs[1:]:
                    run.text = ""
            else:
                para.add_run(text)
            return
    # No paragraphs — add one
    cell.add_paragraph(text)


def _remove_row(row):
    """Remove a row element from its parent table."""
    tbl = row._tr.getparent()
    tbl.remove(row._tr)


def _append_table_row(table, cell_texts: list[str], font_size: int = 12):
    """Add a new row to the table with given cell values."""
    row = table.add_row()
    for i, text in enumerate(cell_texts):
        if i < len(row.cells):
            for para in row.cells[i].paragraphs:
                if para.runs:
                    para.runs[0].text = text
                    para.runs[0].font.size = Pt(font_size)
                else:
                    run = para.add_run(text)
                    run.font.size = Pt(font_size)
                break
