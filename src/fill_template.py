"""
Fill official TransPak Word templates with test data.

Supports ISTA 2A, 3A, 3B, and Simplified templates (the real .docx files,
not the fake token-based templates used during development).

Strategy:
  1. Replace known placeholder paragraphs (product name, operators, date, revision)
  2. Refill the Equipment table from EQUIPMENT.csv data
  3. Update the Test Sample Description table from SUMMARY.csv
  4. Build a document slot map — every image-capable table, keyed by the section
     heading that precedes it
  5. Match chart files to slots by filename keywords; insert into correct sections
  6. Insert pre-test, accelerometer, and post-test photos into their slots

Chart naming convention (export from Lansmont and name the PNG accordingly):
  CHART_ACCEL1_*.png         → Accelerometer #1 transmissibility slot
  CHART_ACCEL2_*.png         → Accelerometer #2 transmissibility slot
  CHART_SEQ2_*.png or CHART_TIP_*.png      → Test Sequence 2 image slot
  CHART_SEQ3_*.png or CHART_ROT_DROP_*.png → Test Sequence 3 image slot
  CHART_SEQ4_*.png or CHART_INCLINE_*.png  → Test Sequence 4 image slot
  CHART_SEQ5_*.png or CHART_VIB_*.png      → Test Sequence 5 (vibration) slot
  CHART_SEQ6_*.png  → Test Sequence 6 slot
  … and so on through SEQ11.
  CHART_FACE3_*.png, CHART_FACE1_*.png … → 3A/2A transmissibility slots by face label
"""

import logging
import re
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
    # Operator lines (all templates)
    "Test Operator #1 Name | Email":            "operator_1",
    " \tTest Operator #1 Name | Email ":        "operator_1",
    "Test Operator #2 Name | Email":            "operator_2",
    "\tTest Operator #2 Name | Email":          "operator_2",
    "Test Operator #3 Name | Email":            "operator_3",
    "Jack Parent | jack.parent@transpak.com":   "operator_1",
    "Annie Schaubel | annie.schaubel@transpak.com": "operator_2",
    # Revision
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

# Chart filename keyword (uppercase) → section heading in template
# More specific keywords listed first — matched in order.
_CHART_SECTION_KEYWORDS: list[tuple[str, str]] = [
    # Transmissibility by accelerometer (3B)
    ("ACCEL1",          "Accelerometer #1 Transmissibility"),
    ("ACCEL_1",         "Accelerometer #1 Transmissibility"),
    ("ACCEL2",          "Accelerometer #2 Transmissibility"),
    ("ACCEL_2",         "Accelerometer #2 Transmissibility"),
    # Transmissibility by face (3A / 2A)
    ("FACE3_OTR",       "Face 3 Over-The-Road"),
    ("FACE3_PUD",       "Face 3 Pick Up"),
    ("FACE3",           "Face 3"),
    ("FACE1",           "Face 1"),
    ("FACE4",           "Face 4"),
    ("FACE6",           "Face 6"),
    # Sequence-numbered charts
    ("SEQ2",            "Test Sequence 2"),
    ("SEQ3",            "Test Sequence 3"),
    ("SEQ4",            "Test Sequence 4"),
    ("SEQ5",            "Test Sequence 5"),
    ("SEQ6",            "Test Sequence 6"),
    ("SEQ7",            "Test Sequence 7"),
    ("SEQ8",            "Test Sequence 8"),
    ("SEQ9",            "Test Sequence 9"),
    ("SEQ10",           "Test Sequence 10"),
    ("SEQ11",           "Test Sequence 11"),
    # Keyword aliases (when file not numbered by sequence)
    ("TIP_OVER",        "Test Sequence 2"),
    ("TIPOVER",         "Test Sequence 2"),
    ("ROT_DROP_1",      "Test Sequence 3"),
    ("ROTDROP1",        "Test Sequence 3"),
    ("INCLINE_1",       "Test Sequence 4"),
    ("INCLINE1",        "Test Sequence 4"),
    ("VIB",             "Test Sequence 5"),
    ("VIBRATION",       "Test Sequence 5"),
    ("FORK",            "Test Sequence 6"),
    ("FLAT_PUSH",       "Test Sequence 6"),
    ("ELEVATED_PUSH",   "Test Sequence 7"),
    ("ELEVATED_ROT",    "Test Sequence 8"),
    ("HANDLING",        "Test Sequence 9"),
    ("ROT_DROP_2",      "Test Sequence 10"),
    ("ROTDROP2",        "Test Sequence 10"),
    ("INCLINE_2",       "Test Sequence 11"),
    ("INCLINE2",        "Test Sequence 11"),
    # Pre-test / Post-test
    ("PRETEST",         "Pretest Preparation"),
    ("PRE_TEST",        "Pretest Preparation"),
    ("POSTTEST",        "Post-Test Photos"),
    ("POST_TEST",       "Post-Test Photos"),
]


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
    Returns a list of inserted-image info dicts.
    """
    inserted: list[dict] = []

    _fill_header_paragraphs(doc, parsed)
    _fill_equipment_table(doc, parsed.get("_equipment_rows", []))
    _fill_sample_description(doc, parsed)

    photo_cfg = cfg["photo_settings"]
    chart_cfg = cfg["chart_settings"]
    max_pw, max_ph = photo_cfg["max_width_inches"], photo_cfg["max_height_inches"]
    max_cw, max_ch = chart_cfg["max_width_inches"], chart_cfg["max_height_inches"]

    # Build the document's image slot map once (section heading → tables)
    slot_map = _build_slot_map(doc)
    logger.debug("Slot map sections: %s", list(slot_map.keys()))

    # ── Charts: match by filename keyword → section → insert ──────────────────
    charts = list(discovered.get("charts", []))
    inserted += _insert_charts(slot_map, charts, max_cw, max_ch)

    # ── Accelerometer photos (named tables: "Accelerometer #1 / #2") ──────────
    accel_photos = discovered.get("accel_photos", [])
    for n, keyword in [(0, "Accelerometer #1"), (1, "Accelerometer #2")]:
        table = _find_table_by_first_cell(doc, keyword)
        if table and n < len(accel_photos):
            result = _insert_image_into_row(table, row_idx=1, col_idx=0,
                                            path=accel_photos[n], max_w=max_pw, max_h=max_ph)
            if result:
                result.update({"kind": "IMG", "slot": f"accel_{n+1}"})
                inserted.append(result)

    # ── Pre-test photos ────────────────────────────────────────────────────────
    pre_table = (_find_table_by_first_cell(doc, "Pretest Photos")
                 or _first_table_in_slot(slot_map, "Pretest Preparation")
                 or _find_table_after_heading(doc, "Pre-test Photos"))
    if pre_table and discovered.get("pre_test_photos"):
        inserted += _fill_photo_table(pre_table, discovered["pre_test_photos"],
                                      max_pw, max_ph, "pre", skip_rows=1)

    # ── Sequence photos ────────────────────────────────────────────────────────
    seq_heading_map = {
        "Seq2_Tip_Over":   "Test Sequence 2",
        "Seq3_Rot_Drop_1": "Test Sequence 3",
        "Seq4_Incline_1":  "Test Sequence 4",
        "Seq6_Rot_Drop_2": "Test Sequence 6",
        "Seq7_Incline_2":  "Test Sequence 7",
    }
    for seq_key, photos_list in discovered.get("seq_photos", {}).items():
        if not photos_list:
            continue
        heading = seq_heading_map.get(seq_key)
        if heading:
            table = _first_empty_table_in_slot(slot_map, heading)
            if table:
                inserted += _fill_photo_table(table, photos_list,
                                              max_pw, max_ph, seq_key, skip_rows=0)

    # ── Post-test photos ───────────────────────────────────────────────────────
    post_table = (_find_table_after_heading(doc, "Post-Test Photos")
                  or _first_table_in_slot(slot_map, "Post-Test Photos"))
    if post_table and discovered.get("post_test_photos"):
        inserted += _fill_photo_table(post_table, discovered["post_test_photos"],
                                      max_pw, max_ph, "post", skip_rows=0)

    return inserted


# ─── Chart slot engine ────────────────────────────────────────────────────────

def _build_slot_map(doc: Document) -> dict[str, list]:
    """
    Scan doc body elements in order and build:
      section_heading_text → [tables that can hold images in that section]

    A table is an "image slot" when:
      - Its header row contains "Transmissibility", "Face N", etc. (named chart slot), OR
      - All cells contain no significant text (empty photo/chart slot)
    """
    slot_map: dict[str, list] = {}
    current_section = "_cover"
    slot_map[current_section] = []

    for kind, elem in _get_body_elements(doc):
        if kind == "para":
            text = elem.text.strip()
            if _is_section_heading(text):
                current_section = text
                slot_map.setdefault(current_section, [])
        elif kind == "table":
            if _table_is_image_slot(elem):
                slot_map.setdefault(current_section, []).append(elem)

    return slot_map


def _is_section_heading(text: str) -> bool:
    """True for paragraph texts that define section context for chart/photo routing."""
    if re.match(r"Test Sequence \d+", text):
        return True
    return text in (
        "Pretest Preparation", "Pre-test Photos", "Pretest Photos",
        "Post-Test Photos", "Post Test Photos",
        "Objective", "Test Protocol", "Test Equipment",
        "Test Sample Description",
    )


def _table_is_image_slot(table) -> bool:
    """True when the table is likely a photo or chart container."""
    if not table.rows:
        return False
    header = table.rows[0].cells[0].text.strip()

    # Named transmissibility / chart tables
    if any(kw in header for kw in ("Transmissibility", "Accelerometer #1", "Accelerometer #2")):
        return True
    if re.match(r"Face \d", header):   # "Face 3 Over-The-Road…" etc.
        return True

    # Empty tables — no cell has more than 30 chars of text
    all_text = " ".join(c.text.strip() for row in table.rows for c in row.cells)
    return len(all_text) <= 30 * len(table.rows)


def _chart_section_key(chart_path: Path) -> str | None:
    """
    Map a chart filename to a section heading using _CHART_SECTION_KEYWORDS.
    Returns the heading string or None if no match.
    """
    name_up = chart_path.stem.upper()
    for keyword, section in _CHART_SECTION_KEYWORDS:
        if keyword in name_up:
            return section
    return None


def _insert_charts(
    slot_map: dict[str, list],
    charts: list[Path],
    max_w: float,
    max_h: float,
) -> list[dict]:
    """
    Match each chart file to its section slot and insert it.
    Named transmissibility tables are matched first by header text;
    other charts are matched to the first available empty slot in their section.
    """
    inserted: list[dict] = []
    # Track how many images have been inserted per table (to pick the right cell)
    table_fill_counts: dict[int, int] = {}

    for chart in sorted(charts, key=lambda p: p.name):
        section = _chart_section_key(chart)
        if section is None:
            logger.info("Chart %s: no section keyword match — skipping auto-insert.", chart.name)
            continue

        # Find which table in that section this chart should go into
        candidates = slot_map.get(section, [])
        # For named transmissibility slots, prefer an exact header match
        target = _find_named_slot(candidates, section) or _next_available_table(candidates, table_fill_counts)

        if target is None:
            logger.warning("Chart %s: section '%s' found but no free slot.", chart.name, section)
            continue

        tbl_id = id(target)
        col_idx = table_fill_counts.get(tbl_id, 0)
        result = _insert_image_into_row(target, row_idx=1, col_idx=col_idx, path=chart,
                                        max_w=max_w, max_h=max_h)
        if result:
            table_fill_counts[tbl_id] = col_idx + 1
            result.update({"kind": "CHART", "slot": f"{section}_{chart.stem}"})
            inserted.append(result)
            logger.info("Chart inserted: %s → %s (col %d)", chart.name, section, col_idx)

    return inserted


def _find_named_slot(tables: list, section: str):
    """Within a list of tables, find one whose header matches the section string."""
    for t in tables:
        if not t.rows:
            continue
        header = t.rows[0].cells[0].text.strip()
        if section.lower() in header.lower() or header.lower() in section.lower():
            return t
    return None


def _next_available_table(tables: list, fill_counts: dict):
    """Return the first table that still has capacity (< num_unique_cells images)."""
    for t in tables:
        tbl_id = id(t)
        n_filled = fill_counts.get(tbl_id, 0)
        n_cells  = _count_unique_cells(t, row_idx=1)
        if n_filled < n_cells:
            return t
    return None


def _count_unique_cells(table, row_idx: int) -> int:
    if row_idx >= len(table.rows):
        return 0
    seen: set[int] = set()
    count = 0
    for cell in table.rows[row_idx].cells:
        tc_id = id(cell._tc)
        if tc_id not in seen:
            seen.add(tc_id)
            count += 1
    return count


# ─── Header paragraph replacement ─────────────────────────────────────────────

def _fill_header_paragraphs(doc: Document, parsed: dict):
    """Replace known placeholder paragraphs in the cover/header area."""
    for para in doc.paragraphs:
        raw  = para.text
        text = raw.strip()

        if "Test Date:" in text:
            date_val = parsed.get("test_date", "")
            if date_val and date_val not in raw:
                run = para.add_run(f"   {date_val}")
                if para.runs:
                    try:
                        run.font.size = para.runs[0].font.size
                    except Exception:
                        pass
            continue

        if any(text.startswith(pfx) for pfx in _OBJECTIVE_STARTS):
            objective = parsed.get("objective", "")
            if objective:
                _set_para_text(para, objective)
            continue

        if text in _PARA_REPLACEMENTS or raw in _PARA_REPLACEMENTS:
            field = _PARA_REPLACEMENTS.get(text) or _PARA_REPLACEMENTS.get(raw)
            value = parsed.get(field, "")
            if value:
                _set_para_text(para, value)


def _set_para_text(para, new_text: str):
    if para.runs:
        para.runs[0].text = new_text
        for run in para.runs[1:]:
            run.text = ""
    else:
        para.add_run(new_text)


# ─── Equipment table ──────────────────────────────────────────────────────────

def _fill_equipment_table(doc: Document, equipment_rows: list[dict]):
    if not equipment_rows:
        return
    table = _find_table_by_first_cell(doc, "Equipment Used")
    if not table or len(table.rows) < 2:
        logger.warning("Equipment table not found in template.")
        return

    col_headers = [c.text.strip().lower() for c in table.rows[1].cells]
    for row in list(table.rows)[2:]:
        _remove_row(row)

    for eq in equipment_rows:
        values = [eq.get(_EQUIP_COL_MAP.get(col, col), "") for col in col_headers]
        _append_table_row(table, values, font_size=12)

    logger.info("Equipment table filled: %d row(s).", len(equipment_rows))


# ─── Test Sample Description table ───────────────────────────────────────────

def _fill_sample_description(doc: Document, parsed: dict):
    table = (_find_table_by_first_cell(doc, "External Packaging")
             or _find_table_by_first_cell(doc, "Packaged-Product Materials"))
    if not table or len(table.rows) < 2:
        return

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

    seen: set[int] = set()
    for cell, col_h in zip(table.rows[1].cells, col_headers):
        tc_id = id(cell._tc)
        if tc_id in seen:
            continue
        seen.add(tc_id)
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
    inserted = []
    photo_iter = iter(photos)
    seen: set[int] = set()

    for row_idx, row in enumerate(table.rows):
        if row_idx < skip_rows:
            continue
        cell_texts = [c.text.strip() for c in row.cells]
        if all(len(t) > 15 for t in cell_texts if t):
            continue  # caption row — skip

        for cell in row.cells:
            tc_id = id(cell._tc)
            if tc_id in seen:
                continue
            seen.add(tc_id)

            photo = next(photo_iter, None)
            if photo is None:
                return inserted

            result = _insert_image_in_cell(cell, photo, max_w, max_h)
            if result:
                result.update({"kind": "IMG", "slot": f"{slot_prefix}_{len(inserted)+1}"})
                inserted.append(result)

    return inserted


def _insert_image_into_row(table, row_idx: int, col_idx: int,
                            path: Path, max_w: float, max_h: float) -> dict | None:
    if row_idx >= len(table.rows):
        return None
    row = table.rows[row_idx]
    seen: set[int] = set()
    real_col = 0
    for cell in row.cells:
        tc_id = id(cell._tc)
        if tc_id in seen:
            continue
        seen.add(tc_id)
        if real_col == col_idx:
            return _insert_image_in_cell(cell, path, max_w, max_h)
        real_col += 1
    return None


# ─── Document body helpers ────────────────────────────────────────────────────

def _get_body_elements(doc: Document) -> list[tuple[str, object]]:
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


def _find_table_by_first_cell(doc: Document, keyword: str):
    for table in doc.tables:
        if not table.rows:
            continue
        first = table.rows[0].cells[0].text.strip()
        if first == keyword or first.startswith(keyword):
            return table
    return None


def _find_table_after_heading(doc: Document, heading_text: str):
    found = False
    for kind, elem in _get_body_elements(doc):
        if kind == "para" and elem.text.strip() == heading_text:
            found = True
        elif kind == "table" and found:
            return elem
    return None


def _first_table_in_slot(slot_map: dict, section: str):
    tables = slot_map.get(section, [])
    return tables[0] if tables else None


def _first_empty_table_in_slot(slot_map: dict, section: str):
    for t in slot_map.get(section, []):
        all_text = " ".join(c.text.strip() for row in t.rows for c in row.cells)
        if len(all_text) < 30:
            return t
    return None


# ─── Cell helpers ─────────────────────────────────────────────────────────────

def _insert_image_in_cell(cell, image_path: Path, max_w: float, max_h: float) -> dict | None:
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
    for p_elem in tc.findall(qn("w:p")):
        for r_elem in list(p_elem.findall(qn("w:r"))):
            p_elem.remove(r_elem)
        for drawing in list(p_elem.findall(".//" + qn("w:drawing"))):
            drawing.getparent().remove(drawing)
    all_paras = tc.findall(qn("w:p"))
    for p_elem in all_paras[1:]:
        tc.remove(p_elem)


def _set_cell_text(cell, text: str):
    for i, para in enumerate(cell.paragraphs):
        if i == 0:
            if para.runs:
                para.runs[0].text = text
                for run in para.runs[1:]:
                    run.text = ""
            else:
                para.add_run(text)
            return
    cell.add_paragraph(text)


def _remove_row(row):
    tbl = row._tr.getparent()
    tbl.remove(row._tr)


def _append_table_row(table, cell_texts: list[str], font_size: int = 12):
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
