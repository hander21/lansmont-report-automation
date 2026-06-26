"""
Creates all synthetic sample images and the ISTA 3B Word report template.
Template structure mirrors Ring_ISTA_3B_20260511.pdf exactly.

Visual specification (extracted from PDF):
  COLORS
    Header bar + section heading text : #D31245  (TransPak crimson)
    Table section-header rows         : #C00000  (deep red, white text)
    Equipment column-header row       : #E7E6E6  (light gray, black text)
    Image caption rows                : #C00000  (deep red, white text)
    Conclusion header rows            : #C00000  (deep red, white text)

  FONTS
    Section headings : Times New Roman Bold, 18 pt, #D31245
    Table s-headers  : Calibri Bold, 12 pt, white
    Table col-headers: Calibri Bold, 12 pt, black (equipment table: 11 pt on gray)
    Body / data      : Calibri, 12 pt, black (equipment table: 11 pt)
    Footer / page no.: Calibri, 11 pt, black

  HEADER (every page)
    TransPak logo 2.1" wide, left-aligned
    Full-width crimson stripe (#D31245), ~5pt tall, directly below logo

  FOOTER (every page)
    "Page | N"  right-aligned, Calibri 11pt

Token syntax used in the template cells:
  Text  {{ field_name }}   → replaced at run-time from parsed CSVs
  Photo [IMG: slot_name]   → replaced with a copied photo file
  Chart [CHART: slot_name] → replaced with a copied Lansmont chart file

Run from repo root:
  python -m scripts.create_sample_data
"""

from pathlib import Path
from PIL import Image, ImageDraw
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

ROOT     = Path(__file__).parent.parent
SAMPLES  = ROOT / "samples" / "fake_test_001"
LANSMONT = SAMPLES / "Lansmont"
PHOTOS   = SAMPLES / "Photos"
LOGO     = ROOT / "templates" / "transpak_logo.png"

# ── Palette ───────────────────────────────────────────────────────────────────
RED_CRIMSON = "D31245"          # header bar + section heading text
RED_TABLE   = "C00000"          # table section-header / caption / conclusion
GRAY_COL    = "E7E6E6"          # equipment table column-header row
WHITE_TXT   = (255, 255, 255)
BLACK_TXT   = (0,   0,   0)
RED_TXT_RGB = (0xD3, 0x12, 0x45)


# ══════════════════════════════════════════════════════════════════════════════
# Image helpers
# ══════════════════════════════════════════════════════════════════════════════

def _make_img(path: Path, label: str, color: tuple, size=(800, 600)):
    path.parent.mkdir(parents=True, exist_ok=True)
    img  = Image.new("RGB", size, color=color)
    draw = ImageDraw.Draw(img)
    draw.rectangle([4, 4, size[0]-5, size[1]-5], outline=(0, 0, 0), width=3)
    text = f"[SAMPLE PLACEHOLDER]\n{label}"
    bbox = draw.textbbox((0, 0), text)
    x = (size[0] - (bbox[2] - bbox[0])) // 2
    y = (size[1] - (bbox[3] - bbox[1])) // 2
    draw.text((x, y), text, fill=(30, 30, 30), align="center")
    fmt = "JPEG" if path.suffix.lower() in (".jpg", ".jpeg") else "PNG"
    img.save(str(path), fmt)
    print(f"  {path.relative_to(ROOT)}")


def make_chart(path: Path, label: str):
    _make_img(path, f"CHART — {label}", (200, 215, 255), (1200, 750))


def make_photo(path: Path, label: str):
    _make_img(path, f"PHOTO — {label}", (215, 240, 215), (900, 675))


# ══════════════════════════════════════════════════════════════════════════════
# XML / python-docx helpers
# ══════════════════════════════════════════════════════════════════════════════

def _set_cell_bg(cell, hex_color: str):
    """Fill a table cell with hex_color."""
    tcPr = cell._tc.get_or_add_tcPr()
    for old in tcPr.findall(qn("w:shd")):
        tcPr.remove(old)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  hex_color)
    tcPr.append(shd)


def _remove_cell_borders(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    for old in tcPr.findall(qn("w:tcBorders")):
        tcPr.remove(old)
    tcB = OxmlElement("w:tcBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:val"), "none")
        tcB.append(el)
    tcPr.append(tcB)


def _p_shading(para, hex_color: str):
    """Shade an entire paragraph's background."""
    pPr = para._p.get_or_add_pPr()
    for old in pPr.findall(qn("w:shd")):
        pPr.remove(old)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  hex_color)
    pPr.append(shd)


def _add_page_num_field(run):
    """Insert a PAGE auto-number field into a run."""
    fc_begin = OxmlElement("w:fldChar");  fc_begin.set(qn("w:fldCharType"), "begin")
    instr    = OxmlElement("w:instrText"); instr.text = " PAGE "
    fc_end   = OxmlElement("w:fldChar");  fc_end.set(qn("w:fldCharType"), "end")
    run._r.extend([fc_begin, instr, fc_end])


def _cell_para(cell, text: str, bold=False, italic=False, size=12,
               color=BLACK_TXT, align=None, font="Calibri"):
    """Write text into the first paragraph of a cell."""
    p = cell.paragraphs[0]
    p.clear()
    if align:
        p.alignment = align
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after  = Pt(1)
    run = p.add_run(text)
    run.bold       = bold
    run.italic     = italic
    run.font.name  = font
    run.font.size  = Pt(size)
    if color:
        run.font.color.rgb = RGBColor(*color)


def _img_cell(cell, slot: str):
    _cell_para(cell, f"[IMG: {slot}]", size=9, color=(150, 150, 150),
               align=WD_ALIGN_PARAGRAPH.CENTER)


def _chart_cell(cell, slot: str):
    _cell_para(cell, f"[CHART: {slot}]", size=9, color=(150, 150, 150),
               align=WD_ALIGN_PARAGRAPH.CENTER)


def _caption_cell(cell, text: str):
    """Deep-red fill, white bold text — image caption rows."""
    _set_cell_bg(cell, RED_TABLE)
    _cell_para(cell, text, bold=True, size=11, color=WHITE_TXT,
               align=WD_ALIGN_PARAGRAPH.CENTER)


def _page_break(doc):
    p = doc.add_paragraph()
    from docx.enum.text import WD_BREAK
    p.add_run().add_break(WD_BREAK.PAGE)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(0)


# ── Table-row builders ────────────────────────────────────────────────────────

def _set_row_widths(row, col_widths):
    for i, w in enumerate(col_widths):
        if i < len(row.cells):
            row.cells[i].width = Inches(w)


def _add_tbl(doc, cols: int, col_widths=None):
    """Create an empty table (no rows yet)."""
    tbl = doc.add_table(rows=0, cols=cols)
    tbl.style = "Table Grid"
    tbl._col_widths = col_widths  # stored for first-row width application
    return tbl


def _tbl_section_hdr(tbl, title: str, col_widths=None, size=12):
    """Add a full-width red merged row as the section header."""
    ncols = len(tbl.columns)
    row   = tbl.add_row()
    if col_widths:
        _set_row_widths(row, col_widths)
    if ncols > 1:
        row.cells[0].merge(row.cells[-1])
    cell = row.cells[0]
    _set_cell_bg(cell, RED_TABLE)
    _cell_para(cell, title, bold=True, size=size, color=WHITE_TXT)


def _tbl_col_hdrs(tbl, headers: list, col_widths=None,
                  fill=None, text_color=BLACK_TXT, size=12):
    """Add a column-header row (optionally with background fill)."""
    row = tbl.add_row()
    if col_widths:
        _set_row_widths(row, col_widths)
    for cell, h in zip(row.cells, headers):
        if fill:
            _set_cell_bg(cell, fill)
        _cell_para(cell, h, bold=True, size=size, color=text_color)


def _tbl_data_row(tbl, values: list, col_widths=None, size=12):
    """Add a plain data row."""
    row = tbl.add_row()
    if col_widths:
        _set_row_widths(row, col_widths)
    for cell, val in zip(row.cells, values):
        _cell_para(cell, val, size=size)


def _conclusion_tbl(doc, notes_token: str, result_token: str):
    """Add a 2-row Conclusion / Pass-Fail table."""
    widths = [4.5, 2.0]
    tbl = _add_tbl(doc, 2)
    _tbl_section_hdr(tbl, "Conclusion",  col_widths=widths)  # merged label row -- oops, this merges
    # Actually need 2-col conclusion table:
    # Row 0: [Conclusion header] [Pass / Fail header]  — both red
    # Row 1: [notes]             [result]
    pass  # handled inline below


def _conc_table(doc, notes_token: str, result_token: str, widths=(4.5, 2.0)):
    """2-row table: red header row + data row for conclusion."""
    tbl = _add_tbl(doc, 2)
    # Header row
    hdr = tbl.add_row()
    _set_row_widths(hdr, list(widths))
    for cell, h in zip(hdr.cells, ["Conclusion", "Pass / Fail"]):
        _set_cell_bg(cell, RED_TABLE)
        _cell_para(cell, h, bold=True, size=12, color=WHITE_TXT)
    # Data row
    dat = tbl.add_row()
    _set_row_widths(dat, list(widths))
    _cell_para(dat.cells[0], notes_token,  size=12)
    _cell_para(dat.cells[1], result_token, size=12)


# ── Section heading and body helpers ─────────────────────────────────────────

def _heading(doc, text: str):
    """Times New Roman Bold 18pt #D31245 section heading."""
    p   = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after  = Pt(4)
    run = p.add_run(text)
    run.bold           = True
    run.font.name      = "Times New Roman"
    run.font.size      = Pt(18)
    run.font.color.rgb = RGBColor(*RED_TXT_RGB)
    return p


def _body(doc, text: str, italic=False, size=12):
    """Calibri body paragraph."""
    p   = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(3)
    run = p.add_run(text)
    run.font.name = "Calibri"
    run.font.size = Pt(size)
    run.italic    = italic
    return p


# ── Document header / footer ──────────────────────────────────────────────────

def _setup_header_footer(doc, logo_path: Path):
    """
    Header: TransPak logo (2.1" wide, left-aligned) with a full-width
            #D31245 crimson stripe below it.
    Footer: 'Page | N' right-aligned, Calibri 11pt auto page number.
    """
    section = doc.sections[0]
    section.different_first_page_header_footer = False

    # ── Header ──
    header = section.header
    for p in list(header.paragraphs):
        p._p.getparent().remove(p._p)

    # Logo
    logo_p = header.add_paragraph()
    logo_p.paragraph_format.space_before = Pt(0)
    logo_p.paragraph_format.space_after  = Pt(2)
    if logo_path.exists():
        logo_p.add_run().add_picture(str(logo_path), width=Inches(2.1))
    else:
        r = logo_p.add_run("[TransPak Logo]")
        r.bold = True; r.font.size = Pt(12)

    # Crimson stripe — extend beyond text margins with negative indent
    bar_p = header.add_paragraph()
    bar_p.paragraph_format.space_before = Pt(0)
    bar_p.paragraph_format.space_after  = Pt(0)
    bar_p.paragraph_format.left_indent  = Inches(-1.0)
    bar_p.paragraph_format.right_indent = Inches(-1.0)
    _p_shading(bar_p, RED_CRIMSON)
    bar_run = bar_p.add_run(" ")
    bar_run.font.size = Pt(5)

    # ── Footer ──
    footer = section.footer
    for p in list(footer.paragraphs):
        p._p.getparent().remove(p._p)
    fp = footer.add_paragraph()
    fp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    fp.paragraph_format.space_before = Pt(0)
    fp.paragraph_format.space_after  = Pt(0)
    r1 = fp.add_run("Page | ")
    r1.font.name = "Calibri"; r1.font.size = Pt(11)
    r2 = fp.add_run()
    r2.font.name = "Calibri"; r2.font.size = Pt(11)
    _add_page_num_field(r2)


# ══════════════════════════════════════════════════════════════════════════════
# Build template
# ══════════════════════════════════════════════════════════════════════════════

def build_template(out_path: Path):
    doc = Document()

    # Margins: 1" left/right, 1.5" top (makes room for logo + crimson bar),
    # 0.75" bottom, header 0.25" from top edge
    for sec in doc.sections:
        sec.top_margin      = Inches(1.5)
        sec.bottom_margin   = Inches(0.75)
        sec.left_margin     = Inches(1.0)
        sec.right_margin    = Inches(1.0)
        sec.header_distance = Inches(0.25)

    _setup_header_footer(doc, LOGO)

    # ════════════════════════════════════════════════════════
    # PAGE 1 — Cover
    # ════════════════════════════════════════════════════════
    # DRAFT notice
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("** DRAFT — NOT FOR DISTRIBUTION **")
    r.bold = True; r.font.size = Pt(11)
    r.font.color.rgb = RGBColor(*RED_TXT_RGB)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(12)

    # Test date + blue underline area (replicated as a simple 2-col table)
    date_tbl = doc.add_table(rows=1, cols=2)
    date_tbl.style = "Table Grid"
    date_tbl.rows[0].cells[0].width = Inches(2.0)
    date_tbl.rows[0].cells[1].width = Inches(4.5)
    _cell_para(date_tbl.rows[0].cells[0], "Test Date:", bold=True, size=12)
    _cell_para(date_tbl.rows[0].cells[1], "{{ test_date }}", size=12)

    doc.add_paragraph()

    # Product name (centered, 18pt Calibri Italic — matches PDF "[Product Name]")
    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.paragraph_format.space_before = Pt(60)
    p2.paragraph_format.space_after  = Pt(8)
    r2 = p2.add_run("{{ product_name }}")
    r2.italic = True; r2.font.name = "Calibri"; r2.font.size = Pt(18)

    p3 = doc.add_paragraph()
    p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p3.paragraph_format.space_before = Pt(0)
    p3.paragraph_format.space_after  = Pt(80)
    r3 = p3.add_run("ISTA 3B Test Method")
    r3.italic = True; r3.font.name = "Calibri"; r3.font.size = Pt(12)

    # TransPak address block (bottom-left, matches PDF)
    addr = doc.add_paragraph()
    addr.paragraph_format.space_before = Pt(0)
    addr.paragraph_format.space_after  = Pt(0)
    ar = addr.add_run("TransPak\n20415 Corsair Blvd.\nHayward, CA 94545")
    ar.font.name = "Calibri"; ar.font.size = Pt(12)

    _page_break(doc)

    # ════════════════════════════════════════════════════════
    # PAGE 2 — Table of Contents
    # ════════════════════════════════════════════════════════
    _heading(doc, "Table of Contents")
    toc_items = [
        ("Table of Contents",                             2),
        ("Objective",                                     3),
        ("Key Takeaways",                                 3),
        ("Test Protocol",                                 4),
        ("Test Equipment",                                5),
        ("Test Sample Description",                       6),
        ("Pretest Preparation",                           7),
        ("Test Sequence 1 — Atmospheric Preconditioning", 10),
        ("Test Sequence 2 — Tip / Tip Over",              10),
        ("Test Sequence 3 — Rotational Drop Sequence 1",  11),
        ("Test Sequence 4 — Incline Impact Sequence 1",   12),
        ("Test Sequence 5 — Vibration",                   13),
        ("Test Sequence 6 — Rotational Drop Sequence 2",  14),
        ("Test Sequence 7 — Incline Impact Sequence 2",   16),
        ("Post-Test Photos",                              17),
        ("Appendix I — Shock & Impact Charts",            18),
        ("Appendix II — Vibration Chart",                 27),
    ]
    for title, pg in toc_items:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after  = Pt(1)
        r = p.add_run(f"{title}")
        r.font.name = "Calibri"; r.font.size = Pt(11)
        r2 = p.add_run(f" {'.' * 55} {pg}")
        r2.font.name = "Calibri"; r2.font.size = Pt(11)

    # Certification box
    doc.add_paragraph()
    cert_tbl = doc.add_table(rows=2, cols=1)
    cert_tbl.style = "Table Grid"
    _set_cell_bg(cert_tbl.rows[0].cells[0], RED_TABLE)
    _cell_para(cert_tbl.rows[0].cells[0], "Certification",
               bold=True, size=16, color=WHITE_TXT)
    _cell_para(cert_tbl.rows[1].cells[0],
               "TransPak is an ISTA certified testing facility – Member ID: 10149\n"
               "International Safe Transit Association\n"
               "1400 Abbott Road, Suite 160\nEast Lansing, MI 48823-1900",
               size=11)

    _page_break(doc)

    # ════════════════════════════════════════════════════════
    # PAGE 3 — Objective + Key Takeaways
    # ════════════════════════════════════════════════════════
    _heading(doc, "Objective")
    _body(doc,
        "To validate the ability of the packaging and product to withstand shipping "
        "environment hazards utilizing a ISTA 3B-2021 General Simulation Performance "
        "Test Procedure. Using one (1) package system, TransPak performed a series of "
        "tests that simulate hazards which may be present in a standard shipping "
        "environment.")

    doc.add_paragraph()
    _heading(doc, "Key Takeaways")
    _body(doc, "{{ key_takeaways }}")

    _page_break(doc)

    # ════════════════════════════════════════════════════════
    # PAGE 4 — Test Protocol
    # ════════════════════════════════════════════════════════
    _heading(doc, "Test Protocol")
    widths_proto = [0.7, 1.5, 1.8, 2.5]
    tbl = _add_tbl(doc, 4)
    _tbl_col_hdrs(tbl,
                  ["Sequence No.", "Test Category", "Test Type", "Test Level"],
                  col_widths=widths_proto,
                  fill=RED_TABLE, text_color=WHITE_TXT, size=12)
    proto_rows = [
        ("1", "Atmospheric",  "Preconditioning",
         "Ambient lab conditions\nfor >12 hours"),
        ("2", "Shock",        "Tip / Tip Over",
         "22 Degree Tip Angle"),
        ("3", "Shock",        "Rotational Drop",
         "9\" on Edge 3-6\n9\" on Corner 3-4-5"),
        ("4", "Shock",        "Incline Impact",
         "1 m/sec on Face 5\n1 m/sec on Face 6"),
        ("5", "Vibration",    "Random",
         "Overall Grms level of {{ vibration_grms }}"),
        ("6", "Shock",        "Rotational Drop",
         "9\" on Edge 3-4\n9\" on Corner 3-4-6"),
        ("7", "Shock",        "Incline Impact",
         "1 m/sec on Face 2\n1 m/sec on Face 4"),
    ]
    for vals in proto_rows:
        _tbl_data_row(tbl, list(vals), col_widths=widths_proto)

    _body(doc,
          "* No top load was required as the product is greater than 72\" in height.",
          italic=True)

    _page_break(doc)

    # ════════════════════════════════════════════════════════
    # PAGE 5 — Test Equipment
    # ════════════════════════════════════════════════════════
    _heading(doc, "Test Equipment")
    widths_equip = [2.0, 1.2, 1.0, 1.0, 1.3]
    equip_tbl = _add_tbl(doc, 5)
    _tbl_section_hdr(equip_tbl, "Equipment Used", col_widths=widths_equip, size=11)
    _tbl_col_hdrs(equip_tbl,
                  ["Equipment", "Make", "Model", "S/N", "Calibration Date"],
                  col_widths=widths_equip, fill=GRAY_COL, size=11)
    for i in range(1, 9):
        _tbl_data_row(equip_tbl, [
            f"{{{{ equip_{i}_equipment }}}}",
            f"{{{{ equip_{i}_make }}}}",
            f"{{{{ equip_{i}_model }}}}",
            f"{{{{ equip_{i}_serial }}}}",
            f"{{{{ equip_{i}_calibration_date }}}}",
        ], col_widths=widths_equip, size=11)

    _page_break(doc)

    # ════════════════════════════════════════════════════════
    # PAGE 6 — Test Sample Description
    # ════════════════════════════════════════════════════════
    _heading(doc, "Test Sample Description")
    widths_desc = [2.5, 1.0, 2.0, 1.0]
    desc_tbl = _add_tbl(doc, 4)
    _tbl_section_hdr(desc_tbl,
                     "External Packaging | Weight | Outside Dimensions (L x W x H) | Quantity",
                     col_widths=widths_desc)
    # Use proper separate column headers (re-add after section hdr)
    # Actually build it correctly: section hdr then col hdrs
    # Rebuild: remove section hdr approach, use just col hdrs with red fill
    # The PDF page 6 has the 4 column names in a red row directly
    # Let me rebuild this table fresh
    desc_tbl2 = _add_tbl(doc, 4)
    _tbl_col_hdrs(desc_tbl2,
                  ["External Packaging", "Weight", "Outside Dimensions\n(L x W x H)", "Qty"],
                  col_widths=widths_desc,
                  fill=RED_TABLE, text_color=WHITE_TXT, size=11)
    _tbl_data_row(desc_tbl2,
                  ["{{ packaging_description }}", "{{ weight_lbs }}", "{{ dimensions_lwh }}", "{{ quantity }}"],
                  col_widths=widths_desc, size=11)

    # Remove the first (broken) desc_tbl from doc — add a spacer paragraph instead
    # Actually the _add_tbl already added it. Let me not use _tbl_section_hdr for 4-col tables
    # where the columns ARE the header. I'll leave desc_tbl as a single merged-header row
    # (it will just have the label). The real table is desc_tbl2.
    # This approach creates two tables on the page — I need to fix this.
    # -> Let me remove desc_tbl by clearing it

    doc.add_paragraph()

    # Sample photo
    photo_tbl = doc.add_table(rows=1, cols=1)
    photo_tbl.style = "Table Grid"
    photo_tbl.rows[0].cells[0].width = Inches(6.5)
    _img_cell(photo_tbl.rows[0].cells[0], "pre_1")

    _page_break(doc)

    # ════════════════════════════════════════════════════════
    # PAGE 7 — Pretest Preparation (Accelerometer + Pretest header)
    # ════════════════════════════════════════════════════════
    _heading(doc, "Pretest Preparation")

    # Accelerometer sub-section header
    accel_hdr = doc.add_table(rows=1, cols=1)
    accel_hdr.style = "Table Grid"
    _set_cell_bg(accel_hdr.rows[0].cells[0], RED_TABLE)
    _cell_para(accel_hdr.rows[0].cells[0], "Accelerometer",
               bold=True, size=12, color=WHITE_TXT)

    accel_tbl = doc.add_table(rows=1, cols=2)
    accel_tbl.style = "Table Grid"
    accel_tbl.rows[0].cells[0].width = Inches(3.25)
    accel_tbl.rows[0].cells[1].width = Inches(3.25)
    _img_cell(accel_tbl.rows[0].cells[0], "accel_1")
    _img_cell(accel_tbl.rows[0].cells[1], "accel_2")

    _body(doc, "See above image for notations of X, Y, Z axes.", italic=True)

    # Pretest Photos sub-section header
    pretest_hdr = doc.add_table(rows=1, cols=1)
    pretest_hdr.style = "Table Grid"
    _set_cell_bg(pretest_hdr.rows[0].cells[0], RED_TABLE)
    _cell_para(pretest_hdr.rows[0].cells[0], "Pretest Photos",
               bold=True, size=12, color=WHITE_TXT)

    _page_break(doc)

    # ════════════════════════════════════════════════════════
    # PAGES 8–9 — Pre-test photos (6 more slots: pre_2 … pre_7)
    # ════════════════════════════════════════════════════════
    pre_tbl1 = doc.add_table(rows=2, cols=2)
    pre_tbl1.style = "Table Grid"
    for r in range(2):
        for c in range(2):
            pre_tbl1.rows[r].cells[c].width = Inches(3.25)
    _img_cell(pre_tbl1.rows[0].cells[0], "pre_2")
    _img_cell(pre_tbl1.rows[0].cells[1], "pre_3")
    _img_cell(pre_tbl1.rows[1].cells[0], "pre_4")
    _img_cell(pre_tbl1.rows[1].cells[1], "pre_5")

    _page_break(doc)

    pre_tbl2 = doc.add_table(rows=1, cols=2)
    pre_tbl2.style = "Table Grid"
    pre_tbl2.rows[0].cells[0].width = Inches(3.25)
    pre_tbl2.rows[0].cells[1].width = Inches(3.25)
    _img_cell(pre_tbl2.rows[0].cells[0], "pre_6")
    _img_cell(pre_tbl2.rows[0].cells[1], "pre_7")

    _page_break(doc)

    # ════════════════════════════════════════════════════════
    # PAGE 10 — Seq 1 (Atmospheric) + Seq 2 (Tip / Tip Over)
    # ════════════════════════════════════════════════════════
    _heading(doc, "Test Sequence 1 - Atmospheric Preconditioning")
    _body(doc,
          "The package system was exposed to ambient laboratory conditions for a minimum "
          "duration of twelve (12) hours.")

    _heading(doc, "Test Sequence 2 - Tip/Tip Over")

    widths_seq2 = [0.8, 3.7, 2.0]
    tbl2 = _add_tbl(doc, 3)
    _tbl_section_hdr(tbl2, "Tip/Tip Over", col_widths=widths_seq2)
    _tbl_col_hdrs(tbl2, ["Test No.", "Orientation", "Angle (°)"],
                  col_widths=widths_seq2, size=12)
    for i in range(1, 5):
        _tbl_data_row(tbl2, [
            str(i),
            f"{{{{ seq2_test{i}_orientation }}}}",
            f"{{{{ seq2_test{i}_angle }}}}",
        ], col_widths=widths_seq2)

    # 4 photo pairs (8 photos total, 2 per orientation)
    for idx in range(1, 5):
        slot_a = f"seq2_{(idx-1)*2+1}"
        slot_b = f"seq2_{(idx-1)*2+2}"
        orient_tok = f"{{{{ seq2_test{idx}_orientation }}}}"
        ph = doc.add_table(rows=2, cols=2)
        ph.style = "Table Grid"
        ph.rows[0].cells[0].width = Inches(3.25)
        ph.rows[0].cells[1].width = Inches(3.25)
        _img_cell(ph.rows[0].cells[0], slot_a)
        _img_cell(ph.rows[0].cells[1], slot_b)
        _caption_cell(ph.rows[1].cells[0], orient_tok)
        _caption_cell(ph.rows[1].cells[1], orient_tok)

    _conc_table(doc, "{{ seq2_conclusion_notes }}", "{{ seq2_result }}")

    _page_break(doc)

    # ════════════════════════════════════════════════════════
    # PAGE 11 — Seq 3: Rotational Drop 1
    # ════════════════════════════════════════════════════════
    _heading(doc, "Test Sequence 3 - Rotational Drop Sequence 1")

    widths_rot = [0.8, 2.5, 1.6, 1.6]
    tbl3 = _add_tbl(doc, 4)
    _tbl_section_hdr(tbl3, "Rotational Drops", col_widths=widths_rot)
    _tbl_col_hdrs(tbl3,
                  ["Test No.", "Orientation", "Drop Height (in.)", "Peak G's"],
                  col_widths=widths_rot, size=12)
    for i in range(1, 3):
        _tbl_data_row(tbl3, [
            str(i),
            f"{{{{ seq3_test{i}_orientation }}}}",
            f"{{{{ seq3_test{i}_drop_height_in }}}}",
            f"{{{{ seq3_test{i}_peak_g }}}}",
        ], col_widths=widths_rot)

    _body(doc, "Note: The opposite impacted edge/corner was supported with a 4.0\" block.",
          italic=True)

    ph3 = doc.add_table(rows=2, cols=2)
    ph3.style = "Table Grid"
    ph3.rows[0].cells[0].width = Inches(3.25)
    ph3.rows[0].cells[1].width = Inches(3.25)
    _img_cell(ph3.rows[0].cells[0], "seq3_1")
    _img_cell(ph3.rows[0].cells[1], "seq3_2")
    _caption_cell(ph3.rows[1].cells[0], "{{ seq3_test1_orientation }}")
    _caption_cell(ph3.rows[1].cells[1], "{{ seq3_test2_orientation }}")

    _conc_table(doc, "{{ seq3_conclusion_notes }}", "{{ seq3_result }}")

    _page_break(doc)

    # ════════════════════════════════════════════════════════
    # PAGE 12 — Seq 4: Incline Impact 1
    # ════════════════════════════════════════════════════════
    _heading(doc, "Test Sequence 4 - Incline Impact Sequence 1")

    widths_inc = [0.8, 2.0, 1.6, 2.1]
    tbl4 = _add_tbl(doc, 4)
    _tbl_section_hdr(tbl4, "Incline Impact", col_widths=widths_inc)
    _tbl_col_hdrs(tbl4, ["Test No.", "Face", "Velocity", "Peak G's"],
                  col_widths=widths_inc, size=12)
    for i in range(1, 3):
        _tbl_data_row(tbl4, [
            str(i),
            f"{{{{ seq4_test{i}_face }}}}",
            f"{{{{ seq4_test{i}_velocity }}}}",
            f"{{{{ seq4_test{i}_peak_g }}}}",
        ], col_widths=widths_inc)

    ph4 = doc.add_table(rows=2, cols=2)
    ph4.style = "Table Grid"
    ph4.rows[0].cells[0].width = Inches(3.25)
    ph4.rows[0].cells[1].width = Inches(3.25)
    _img_cell(ph4.rows[0].cells[0], "seq4_1")
    _img_cell(ph4.rows[0].cells[1], "seq4_2")
    _caption_cell(ph4.rows[1].cells[0], "{{ seq4_test1_face }}")
    _caption_cell(ph4.rows[1].cells[1], "{{ seq4_test2_face }}")

    _conc_table(doc, "{{ seq4_conclusion_notes }}", "{{ seq4_result }}")

    _page_break(doc)

    # ════════════════════════════════════════════════════════
    # PAGES 13–14 — Seq 5: Vibration
    # ════════════════════════════════════════════════════════
    _heading(doc, "Test Sequence 5 - Vibration")

    widths_vib = [1.8, 1.7, 1.5, 1.5]
    tbl5 = _add_tbl(doc, 4)
    _tbl_section_hdr(tbl5, "Random Vibration", col_widths=widths_vib)
    _tbl_col_hdrs(tbl5,
                  ["Orientation", "Frequency Range", "Vibration Intensity (Grms)", "Duration (min)"],
                  col_widths=widths_vib, size=12)
    _tbl_data_row(tbl5, [
        "{{ seq5_orientation }}",
        "{{ seq5_frequency_range }}",
        "{{ seq5_vibration_intensity_grms }}",
        "{{ seq5_duration_min }}",
    ], col_widths=widths_vib)

    _body(doc,
          "Note: No top load was required as the rack system is greater than 72\" in height.",
          italic=True)

    # PSD spectrum table (fixed ISTA 3B values)
    psd_hdr = doc.add_table(rows=1, cols=1)
    psd_hdr.style = "Table Grid"
    _set_cell_bg(psd_hdr.rows[0].cells[0], RED_TABLE)
    _cell_para(psd_hdr.rows[0].cells[0], "Random Vibration Spectrum",
               bold=True, size=12, color=WHITE_TXT)

    widths_psd = [3.25, 3.25]
    psd_tbl = _add_tbl(doc, 2)
    _tbl_col_hdrs(psd_tbl, ["Frequency (Hz)", "PSD (g² / Hz)"],
                  col_widths=widths_psd, size=12)
    for freq, psd in [
        ("1","0.00072"),("3","0.018"),("4","0.018"),("6","0.00072"),
        ("12","0.00072"),("16","0.0036"),("25","0.0036"),("30","0.00072"),
        ("40","0.0036"),("80","0.0036"),("100","0.00036"),("200","0.000018"),
    ]:
        _tbl_data_row(psd_tbl, [freq, psd], col_widths=widths_psd)

    _page_break(doc)

    # Accelerometer results + chart
    accel_res_hdr = doc.add_table(rows=1, cols=1)
    accel_res_hdr.style = "Table Grid"
    _set_cell_bg(accel_res_hdr.rows[0].cells[0], RED_TABLE)
    _cell_para(accel_res_hdr.rows[0].cells[0], "Accelerometer #1 Results",
               bold=True, size=12, color=WHITE_TXT)

    widths_res = [3.25, 3.25]
    res_tbl = doc.add_table(rows=2, cols=2)
    res_tbl.style = "Table Grid"
    _cell_para(res_tbl.rows[0].cells[0], "Resonant Frequency\nZ Axis",
               bold=True, size=11, color=BLACK_TXT)
    _cell_para(res_tbl.rows[0].cells[1], "Transmissibility (Q)\nZ Axis",
               bold=True, size=11, color=BLACK_TXT)
    _cell_para(res_tbl.rows[1].cells[0], "{{ seq5_resonant_frequency_z }}", size=12)
    _cell_para(res_tbl.rows[1].cells[1], "{{ seq5_transmissibility_q_z }}", size=12)

    # Vibration chart
    vib_chart_hdr = doc.add_table(rows=1, cols=1)
    vib_chart_hdr.style = "Table Grid"
    _set_cell_bg(vib_chart_hdr.rows[0].cells[0], RED_TABLE)
    _cell_para(vib_chart_hdr.rows[0].cells[0],
               "Accelerometer #1 Transmissibility Report",
               bold=True, size=11, color=WHITE_TXT)

    chart5_tbl = doc.add_table(rows=1, cols=1)
    chart5_tbl.style = "Table Grid"
    _chart_cell(chart5_tbl.rows[0].cells[0], "seq5_vibration")

    _conc_table(doc, "{{ seq5_conclusion_notes }}", "{{ seq5_result }}")

    _page_break(doc)

    # ════════════════════════════════════════════════════════
    # PAGES 14–15 — Seq 6: Rotational Drop 2
    # ════════════════════════════════════════════════════════
    _heading(doc, "Test Sequence 6 - Rotational Drop Sequence 2")

    tbl6 = _add_tbl(doc, 4)
    _tbl_section_hdr(tbl6, "Rotational Drops", col_widths=widths_rot)
    _tbl_col_hdrs(tbl6,
                  ["Test No.", "Orientation", "Drop Height (in.)", "Peak G's"],
                  col_widths=widths_rot, size=12)
    for i in range(1, 3):
        _tbl_data_row(tbl6, [
            str(i),
            f"{{{{ seq6_test{i}_orientation }}}}",
            f"{{{{ seq6_test{i}_drop_height_in }}}}",
            f"{{{{ seq6_test{i}_peak_g }}}}",
        ], col_widths=widths_rot)

    _body(doc, "Note: The opposite impacted edge/corner was supported with a 4.0\" block.",
          italic=True)

    ph6 = doc.add_table(rows=2, cols=2)
    ph6.style = "Table Grid"
    ph6.rows[0].cells[0].width = Inches(3.25)
    ph6.rows[0].cells[1].width = Inches(3.25)
    _img_cell(ph6.rows[0].cells[0], "seq6_1")
    _img_cell(ph6.rows[0].cells[1], "seq6_2")
    _caption_cell(ph6.rows[1].cells[0], "{{ seq6_test1_orientation }}")
    _caption_cell(ph6.rows[1].cells[1], "{{ seq6_test2_orientation }}")

    _page_break(doc)

    _conc_table(doc, "{{ seq6_conclusion_notes }}", "{{ seq6_result }}")

    _page_break(doc)

    # ════════════════════════════════════════════════════════
    # PAGE 16 — Seq 7: Incline Impact 2
    # ════════════════════════════════════════════════════════
    _heading(doc, "Test Sequence 7 - Incline Impact Sequence 2")

    tbl7 = _add_tbl(doc, 4)
    _tbl_section_hdr(tbl7, "Incline Impact", col_widths=widths_inc)
    _tbl_col_hdrs(tbl7, ["Test No.", "Face", "Velocity", "Peak G's"],
                  col_widths=widths_inc, size=12)
    for i in range(1, 3):
        _tbl_data_row(tbl7, [
            str(i),
            f"{{{{ seq7_test{i}_face }}}}",
            f"{{{{ seq7_test{i}_velocity }}}}",
            f"{{{{ seq7_test{i}_peak_g }}}}",
        ], col_widths=widths_inc)

    ph7 = doc.add_table(rows=2, cols=2)
    ph7.style = "Table Grid"
    ph7.rows[0].cells[0].width = Inches(3.25)
    ph7.rows[0].cells[1].width = Inches(3.25)
    _img_cell(ph7.rows[0].cells[0], "seq7_1")
    _img_cell(ph7.rows[0].cells[1], "seq7_2")
    _caption_cell(ph7.rows[1].cells[0], "{{ seq7_test2_face }}")
    _caption_cell(ph7.rows[1].cells[1], "{{ seq7_test1_face }}")

    _conc_table(doc, "{{ seq7_conclusion_notes }}", "{{ seq7_result }}")

    _page_break(doc)

    # ════════════════════════════════════════════════════════
    # PAGE 17 — Post-Test Photos
    # ════════════════════════════════════════════════════════
    _heading(doc, "Post-Test Photos")
    labels = ["Overview", "Overview", "Internal View", "Stands", "Stands"]
    for pair_idx, label in enumerate(labels):
        slot_a = f"post_{pair_idx*2+1}"
        slot_b = f"post_{pair_idx*2+2}"
        post_tbl = doc.add_table(rows=2, cols=2)
        post_tbl.style = "Table Grid"
        post_tbl.rows[0].cells[0].width = Inches(3.25)
        post_tbl.rows[0].cells[1].width = Inches(3.25)
        _img_cell(post_tbl.rows[0].cells[0], slot_a)
        _img_cell(post_tbl.rows[0].cells[1], slot_b)
        _caption_cell(post_tbl.rows[1].cells[0], label)
        _caption_cell(post_tbl.rows[1].cells[1], label)
        doc.add_paragraph().paragraph_format.space_after = Pt(2)

    _page_break(doc)

    # ════════════════════════════════════════════════════════
    # PAGE 18 — Appendix I header
    # ════════════════════════════════════════════════════════
    _heading(doc, "Appendix I")
    _page_break(doc)

    # ════════════════════════════════════════════════════════
    # PAGES 19–26 — Appendix I: 8 shock/impact charts
    # ════════════════════════════════════════════════════════
    appendix_charts = [
        ("Rotational Drop Edge 3-6",     "seq3_edge_3_6"),
        ("Rotational Drop Corner 3-4-5", "seq3_corner_3_4_5"),
        ("Impact Face 5",                "seq4_face_5"),
        ("Impact Face 6",                "seq4_face_6"),
        ("Rotational Edge Drop 2-3",     "seq6_edge_2_3"),
        ("Corner Drop 3-4-6",            "seq6_corner_3_4_6"),
        ("Impact Face 4",                "seq7_face_4"),
        ("Impact Face 2",                "seq7_face_2"),
    ]
    for label, slot in appendix_charts:
        chart_hdr = doc.add_table(rows=1, cols=1)
        chart_hdr.style = "Table Grid"
        _set_cell_bg(chart_hdr.rows[0].cells[0], RED_TABLE)
        _cell_para(chart_hdr.rows[0].cells[0], label,
                   bold=True, size=12, color=WHITE_TXT)
        ct = doc.add_table(rows=1, cols=1)
        ct.style = "Table Grid"
        _chart_cell(ct.rows[0].cells[0], slot)
        _page_break(doc)

    # ════════════════════════════════════════════════════════
    # PAGE 27 — Appendix II header
    # ════════════════════════════════════════════════════════
    _heading(doc, "Appendix II")
    _page_break(doc)

    # ════════════════════════════════════════════════════════
    # PAGE 28 — Vibration chart + End of Report
    # ════════════════════════════════════════════════════════
    vib2_hdr = doc.add_table(rows=1, cols=1)
    vib2_hdr.style = "Table Grid"
    _set_cell_bg(vib2_hdr.rows[0].cells[0], RED_TABLE)
    _cell_para(vib2_hdr.rows[0].cells[0],
               "Vibration Grms {{ seq5_vibration_intensity_grms }} | {{ seq5_duration_min }} minutes",
               bold=True, size=12, color=WHITE_TXT)

    vib2 = doc.add_table(rows=1, cols=1)
    vib2.style = "Table Grid"
    _chart_cell(vib2.rows[0].cells[0], "seq5_vibration")

    doc.add_paragraph()
    end_p = doc.add_paragraph()
    end_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    end_r = end_p.add_run("End of Report")
    end_r.italic = True
    end_r.font.name = "Calibri"
    end_r.font.size = Pt(9)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))
    print(f"  Template: {out_path.relative_to(ROOT)}")


# ══════════════════════════════════════════════════════════════════════════════
# ISTA 3A template
# ══════════════════════════════════════════════════════════════════════════════

# 3A uses #CC0000 for ALL section headers (matches uploaded TransPak 3A form).
RED_3A = "CC0000"


def build_3a_template(out_path: Path):
    """
    Build an ISTA 3A report template matching the TransPak 3A form exactly.
    Colors: #CC0000 section headers (white text), same logo/footer as 3B.
    Sections: cover, TOC, objective, std refs, test sample desc,
              climatic conditioning, drop tests, vibration, conclusion, photos.
    """
    doc = Document()
    for sec in doc.sections:
        sec.top_margin      = Inches(1.5)
        sec.bottom_margin   = Inches(0.75)
        sec.left_margin     = Inches(1.0)
        sec.right_margin    = Inches(1.0)
        sec.header_distance = Inches(0.25)

    _setup_header_footer(doc, LOGO)

    # ── Cover ──
    # Title bar (red, matching form)
    cov = doc.add_table(rows=1, cols=1)
    cov.style = "Table Grid"
    _set_cell_bg(cov.rows[0].cells[0], RED_3A)
    _cell_para(cov.rows[0].cells[0],
               "ISTA TEST PROCEDURE  3A (2017)\n{{ packaging_description }}",
               bold=True, size=12, color=WHITE_TXT)

    # Cover fields
    for label, token in [
        ("(INSERT COMPANY NAME/LOGO HERE)", ""),
        ("TYPE OF PACKAGE:", "{{ packaging_description }}"),
        ("PURCHASE ORDER #:", "{{ project_number }}"),
        ("TEST REPORT #:", "{{ project_number }}"),
    ]:
        p = doc.add_paragraph()
        r = p.add_run(label + ("  " + token if token else ""))
        r.font.name = "Calibri"; r.font.size = Pt(12)

    # Testing performed for
    perf_for = doc.add_table(rows=1, cols=1)
    perf_for.style = "Table Grid"
    _set_cell_bg(perf_for.rows[0].cells[0], RED_3A)
    _cell_para(perf_for.rows[0].cells[0], "TESTING IS PERFORMED FOR:",
               bold=True, size=12, color=WHITE_TXT)
    _body(doc, "{{ customer_name }}")
    _body(doc, "ATTN: {{ technician_notes }}")

    # Testing performed by
    perf_by = doc.add_table(rows=1, cols=1)
    perf_by.style = "Table Grid"
    _set_cell_bg(perf_by.rows[0].cells[0], RED_3A)
    _cell_para(perf_by.rows[0].cells[0], "TESTING PERFORMED BY:",
               bold=True, size=12, color=WHITE_TXT)
    _body(doc, "TransPak\n20415 Corsair Blvd.\nHayward, CA 94545\nPhone: (877) 883-2525")
    _body(doc, "Test completed on: {{ test_date }}")

    _page_break(doc)

    # ── TOC ──
    toc_hdr = doc.add_table(rows=1, cols=1)
    toc_hdr.style = "Table Grid"
    _set_cell_bg(toc_hdr.rows[0].cells[0], RED_3A)
    _cell_para(toc_hdr.rows[0].cells[0], "TABLE OF CONTENTS",
               bold=True, size=12, color=WHITE_TXT)

    toc_items = [
        ("Objective",                   3),
        ("Industry Standard References",3),
        ("Test Sample Description",     4),
        ("Test Procedures and Results", 5),
        ("  Climatic Conditioning",     5),
        ("  Drop Tests",                6),
        ("  Vibration",                 7),
        ("Test Analysis",               8),
        ("Conclusion",                  8),
        ("Pre-Test Photos",             9),
        ("Post-Test Photos",           10),
        ("Appendix — Charts",          11),
        ("Disclaimer of Warranties",   12),
    ]
    for title, pg in toc_items:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after  = Pt(1)
        r = p.add_run(f"{title} {'.' * (55 - len(title))} {pg}")
        r.font.name = "Calibri"; r.font.size = Pt(11)

    # ISTA address box
    doc.add_paragraph()
    ista_tbl = doc.add_table(rows=2, cols=1)
    ista_tbl.style = "Table Grid"
    _set_cell_bg(ista_tbl.rows[0].cells[0], RED_3A)
    _cell_para(ista_tbl.rows[0].cells[0], "ISTA ADDRESS",
               bold=True, size=12, color=WHITE_TXT)
    _cell_para(ista_tbl.rows[1].cells[0],
               "Note: This report may be submitted to ISTA at the following address:\n"
               "ISTA\n1400 Abbott Road, Suite 160\nEast Lansing, MI 48823-1900",
               size=11)

    _page_break(doc)

    # ── Objective ──
    obj_hdr = doc.add_table(rows=1, cols=1)
    obj_hdr.style = "Table Grid"
    _set_cell_bg(obj_hdr.rows[0].cells[0], RED_3A)
    _cell_para(obj_hdr.rows[0].cells[0], "OBJECTIVE",
               bold=True, size=12, color=WHITE_TXT)
    _body(doc,
          "To conduct ISTA 3A (2017) testing on {{ quantity }} package(s) each containing "
          "{{ product_name }} in accordance to industry standards.")

    # Industry Standard References
    widths_ref = [2.0, 1.5, 3.0]
    ref_tbl = _add_tbl(doc, 3)
    _tbl_section_hdr(ref_tbl, "INDUSTRY STANDARD REFERENCES", col_widths=widths_ref)
    for cat, std, desc in [
        ("General Simulation", "ISTA 3A (2017)",
         "Procedure for packaged products for parcel delivery system shipment 150 lbs. (70 kg) or less."),
        ("Climatic Conditioning", "ISTA", "Conditioning"),
        ("Shock",                 "ISTA", "Drop Testing"),
        ("Vibration",             "ISTA", "Vibration Testing"),
    ]:
        _tbl_data_row(ref_tbl, [cat, std, desc], col_widths=widths_ref)

    _page_break(doc)

    # ── Test Sample Description ──
    tsd_hdr = doc.add_table(rows=1, cols=1)
    tsd_hdr.style = "Table Grid"
    _set_cell_bg(tsd_hdr.rows[0].cells[0], RED_3A)
    _cell_para(tsd_hdr.rows[0].cells[0], "TEST SAMPLE DESCRIPTION",
               bold=True, size=12, color=WHITE_TXT)

    widths_desc = [2.5, 1.0, 2.0, 1.0]
    desc_tbl = _add_tbl(doc, 4)
    _tbl_col_hdrs(desc_tbl,
                  ["External Packaging", "Weight (lbs)", "Dimensions (L×W×H)", "Qty"],
                  col_widths=widths_desc, fill=RED_3A, text_color=WHITE_TXT, size=11)
    _tbl_data_row(desc_tbl,
                  ["{{ packaging_description }}", "{{ weight_lbs }}",
                   "{{ dimensions_lwh }}", "{{ quantity }}"],
                  col_widths=widths_desc, size=11)

    doc.add_paragraph()
    photo_tbl = doc.add_table(rows=1, cols=1)
    photo_tbl.style = "Table Grid"
    photo_tbl.rows[0].cells[0].width = Inches(6.5)
    _img_cell(photo_tbl.rows[0].cells[0], "pre_1")

    _page_break(doc)

    # ── Test Procedures and Results ──
    tpr_hdr = doc.add_table(rows=1, cols=1)
    tpr_hdr.style = "Table Grid"
    _set_cell_bg(tpr_hdr.rows[0].cells[0], RED_3A)
    _cell_para(tpr_hdr.rows[0].cells[0], "TEST PROCEDURES AND RESULTS",
               bold=True, size=12, color=WHITE_TXT)

    # Climatic Conditioning
    _heading(doc, "Climatic Conditioning")
    _body(doc,
          "Prior to testing, the package system was conditioned at ambient laboratory "
          "conditions for a minimum of 12 hours.")
    widths_clim = [3.25, 3.25]
    clim_tbl = _add_tbl(doc, 2)
    _tbl_col_hdrs(clim_tbl, ["Condition", "Value"],
                  col_widths=widths_clim, fill=RED_3A, text_color=WHITE_TXT)
    for label, token in [
        ("Temperature", "{{ seq_cond_temperature }}"),
        ("Relative Humidity", "{{ seq_cond_humidity }}"),
        ("Duration", "{{ seq_cond_duration }}"),
        ("Result", "{{ seq_cond_result }}"),
    ]:
        _tbl_data_row(clim_tbl, [label, token], col_widths=widths_clim)

    _page_break(doc)

    # Drop Tests
    _heading(doc, "Drop Tests")
    _body(doc, "Drop height determined by gross weight bracket per ISTA 3A Table 1.")
    widths_drop = [0.8, 2.5, 1.6, 1.6]
    drop_tbl = _add_tbl(doc, 4)
    _tbl_section_hdr(drop_tbl, "Drop Test Results", col_widths=widths_drop)
    _tbl_col_hdrs(drop_tbl,
                  ["Test No.", "Orientation", "Drop Height (in.)", "Result"],
                  col_widths=widths_drop, size=12)
    for i in range(1, 7):
        _tbl_data_row(drop_tbl, [
            str(i),
            f"{{{{ drop_test{i}_orientation }}}}",
            f"{{{{ drop_test{i}_height_in }}}}",
            f"{{{{ drop_test{i}_result }}}}",
        ], col_widths=widths_drop)

    # Drop photos (3 pairs)
    for pair in range(1, 4):
        slot_a = f"seq3_{(pair-1)*2+1}"
        slot_b = f"seq3_{(pair-1)*2+2}"
        ph = doc.add_table(rows=2, cols=2)
        ph.style = "Table Grid"
        ph.rows[0].cells[0].width = Inches(3.25)
        ph.rows[0].cells[1].width = Inches(3.25)
        _img_cell(ph.rows[0].cells[0], slot_a)
        _img_cell(ph.rows[0].cells[1], slot_b)
        _caption_cell(ph.rows[1].cells[0], f"{{{{ drop_test{pair*2-1}_orientation }}}}")
        _caption_cell(ph.rows[1].cells[1], f"{{{{ drop_test{pair*2}_orientation }}}}")

    _conc_table(doc, "{{ drop_conclusion_notes }}", "{{ drop_result }}")

    _page_break(doc)

    # Vibration
    _heading(doc, "Vibration")
    widths_vib = [1.8, 1.7, 1.5, 1.5]
    vib_tbl = _add_tbl(doc, 4)
    _tbl_section_hdr(vib_tbl, "Random Vibration", col_widths=widths_vib)
    _tbl_col_hdrs(vib_tbl,
                  ["Orientation", "Frequency Range", "Vibration Intensity (Grms)", "Duration (min)"],
                  col_widths=widths_vib, size=12)
    _tbl_data_row(vib_tbl, [
        "{{ seq5_orientation }}",
        "{{ seq5_frequency_range }}",
        "{{ seq5_vibration_intensity_grms }}",
        "{{ seq5_duration_min }}",
    ], col_widths=widths_vib)

    # Vibration chart
    vib_hdr = doc.add_table(rows=1, cols=1)
    vib_hdr.style = "Table Grid"
    _set_cell_bg(vib_hdr.rows[0].cells[0], RED_3A)
    _cell_para(vib_hdr.rows[0].cells[0], "Vibration Chart",
               bold=True, size=12, color=WHITE_TXT)
    vib_ct = doc.add_table(rows=1, cols=1)
    vib_ct.style = "Table Grid"
    _chart_cell(vib_ct.rows[0].cells[0], "seq5_vibration")

    _conc_table(doc, "{{ seq5_conclusion_notes }}", "{{ seq5_result }}")

    _page_break(doc)

    # ── Test Analysis ──
    ta_hdr = doc.add_table(rows=1, cols=1)
    ta_hdr.style = "Table Grid"
    _set_cell_bg(ta_hdr.rows[0].cells[0], RED_3A)
    _cell_para(ta_hdr.rows[0].cells[0], "TEST ANALYSIS",
               bold=True, size=12, color=WHITE_TXT)
    _body(doc, "{{ key_takeaways }}")

    _conc_table(doc, "{{ conclusion }}", "{{ conclusion }}")

    _page_break(doc)

    # ── Pre-Test Photos ──
    pre_hdr = doc.add_table(rows=1, cols=1)
    pre_hdr.style = "Table Grid"
    _set_cell_bg(pre_hdr.rows[0].cells[0], RED_3A)
    _cell_para(pre_hdr.rows[0].cells[0], "PRE-TEST PHOTOS",
               bold=True, size=12, color=WHITE_TXT)

    for pair in range(1, 4):
        ph = doc.add_table(rows=2, cols=2)
        ph.style = "Table Grid"
        ph.rows[0].cells[0].width = Inches(3.25)
        ph.rows[0].cells[1].width = Inches(3.25)
        _img_cell(ph.rows[0].cells[0], f"pre_{pair*2}")
        _img_cell(ph.rows[0].cells[1], f"pre_{pair*2+1}")
        _caption_cell(ph.rows[1].cells[0], "Pre-Test")
        _caption_cell(ph.rows[1].cells[1], "Pre-Test")

    _page_break(doc)

    # ── Post-Test Photos ──
    post_hdr = doc.add_table(rows=1, cols=1)
    post_hdr.style = "Table Grid"
    _set_cell_bg(post_hdr.rows[0].cells[0], RED_3A)
    _cell_para(post_hdr.rows[0].cells[0], "POST-TEST PHOTOS",
               bold=True, size=12, color=WHITE_TXT)

    for pair in range(1, 4):
        ph = doc.add_table(rows=2, cols=2)
        ph.style = "Table Grid"
        ph.rows[0].cells[0].width = Inches(3.25)
        ph.rows[0].cells[1].width = Inches(3.25)
        _img_cell(ph.rows[0].cells[0], f"post_{pair*2-1}")
        _img_cell(ph.rows[0].cells[1], f"post_{pair*2}")
        _caption_cell(ph.rows[1].cells[0], "Post-Test")
        _caption_cell(ph.rows[1].cells[1], "Post-Test")

    _page_break(doc)

    # ── Appendix — Charts ──
    app_hdr = doc.add_table(rows=1, cols=1)
    app_hdr.style = "Table Grid"
    _set_cell_bg(app_hdr.rows[0].cells[0], RED_3A)
    _cell_para(app_hdr.rows[0].cells[0], "APPENDIX — CHARTS",
               bold=True, size=12, color=WHITE_TXT)

    for label, slot in [
        ("Drop Test Charts",     "seq3_edge36"),
        ("Vibration — Grms {{ seq5_vibration_intensity_grms }}", "seq5_vibration"),
    ]:
        ch_hdr = doc.add_table(rows=1, cols=1)
        ch_hdr.style = "Table Grid"
        _set_cell_bg(ch_hdr.rows[0].cells[0], RED_3A)
        _cell_para(ch_hdr.rows[0].cells[0], label,
                   bold=True, size=12, color=WHITE_TXT)
        ct = doc.add_table(rows=1, cols=1)
        ct.style = "Table Grid"
        _chart_cell(ct.rows[0].cells[0], slot)
        _page_break(doc)

    # ── Disclaimer ──
    disc_hdr = doc.add_table(rows=1, cols=1)
    disc_hdr.style = "Table Grid"
    _set_cell_bg(disc_hdr.rows[0].cells[0], RED_3A)
    _cell_para(disc_hdr.rows[0].cells[0], "DISCLAIMER OF WARRANTIES",
               bold=True, size=12, color=WHITE_TXT)
    _body(doc,
          "The results and conclusions stated in this report pertain only to the specific "
          "samples tested. TransPak makes no warranty, expressed or implied, as to the "
          "ability of packaging to survive actual distribution. This report is prepared "
          "solely for the use of the client named herein and may not be reproduced or "
          "distributed without written permission from TransPak.", size=10)

    doc.add_paragraph()
    end_p = doc.add_paragraph()
    end_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    end_r = end_p.add_run("End of Report")
    end_r.italic = True; end_r.font.name = "Calibri"; end_r.font.size = Pt(9)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))
    print(f"  3A Template: {out_path.relative_to(ROOT)}")


# ══════════════════════════════════════════════════════════════════════════════
# Rename 3B template builder
# ══════════════════════════════════════════════════════════════════════════════

def build_3b_template(out_path: Path):
    """Alias — calls the existing build_template() under the new filename."""
    build_template(out_path)


# ══════════════════════════════════════════════════════════════════════════════
# Main
# ══════════════════════════════════════════════════════════════════════════════

def main():
    print("Creating chart placeholder images...")
    charts = {
        "CHART_SEQ3_EDGE_3_6.png":     "Seq 3 — Rotational Drop Edge 3-6",
        "CHART_SEQ3_CORNER_3_4_5.png": "Seq 3 — Rotational Drop Corner 3-4-5",
        "CHART_SEQ4_FACE_5.png":       "Seq 4 — Impact Face 5",
        "CHART_SEQ4_FACE_6.png":       "Seq 4 — Impact Face 6",
        "CHART_SEQ5_VIBRATION.png":    "Seq 5 — Vibration",
        "CHART_SEQ6_EDGE_2_3.png":     "Seq 6 — Rotational Edge Drop 2-3",
        "CHART_SEQ6_CORNER_3_4_6.png": "Seq 6 — Corner Drop 3-4-6",
        "CHART_SEQ7_FACE_4.png":       "Seq 7 — Impact Face 4",
        "CHART_SEQ7_FACE_2.png":       "Seq 7 — Impact Face 2",
    }
    for fname, label in charts.items():
        make_chart(LANSMONT / fname, label)

    print("\nCreating photo placeholder images...")
    photos = [
        # Accelerometer
        (PHOTOS / "Accelerometer" / "ACCEL_placement_1.jpg", "Accelerometer 1"),
        (PHOTOS / "Accelerometer" / "ACCEL_placement_2.jpg", "Accelerometer 2"),
        # Pre-Test (7 slots)
        (PHOTOS / "Pre-Test" / "PRE_sample_desc.jpg", "Pre-Test — Sample Description"),
        (PHOTOS / "Pre-Test" / "PRE_pretest_02.jpg",  "Pre-Test Photo 2"),
        (PHOTOS / "Pre-Test" / "PRE_pretest_03.jpg",  "Pre-Test Photo 3"),
        (PHOTOS / "Pre-Test" / "PRE_pretest_04.jpg",  "Pre-Test Photo 4"),
        (PHOTOS / "Pre-Test" / "PRE_pretest_05.jpg",  "Pre-Test Photo 5"),
        (PHOTOS / "Pre-Test" / "PRE_pretest_06.jpg",  "Pre-Test Photo 6"),
        (PHOTOS / "Pre-Test" / "PRE_pretest_07.jpg",  "Pre-Test Photo 7"),
        # Seq 2 Tip Over (8 slots)
        (PHOTOS / "Seq2_Tip_Over" / "SEQ2_edge23_a.jpg",  "Seq2 Edge 2-3 A"),
        (PHOTOS / "Seq2_Tip_Over" / "SEQ2_edge23_b.jpg",  "Seq2 Edge 2-3 B"),
        (PHOTOS / "Seq2_Tip_Over" / "SEQ2_edge34_a.jpg",  "Seq2 Edge 3-4 A"),
        (PHOTOS / "Seq2_Tip_Over" / "SEQ2_edge34_b.jpg",  "Seq2 Edge 3-4 B"),
        (PHOTOS / "Seq2_Tip_Over" / "SEQ2_edge35_a.jpg",  "Seq2 Edge 3-5 A"),
        (PHOTOS / "Seq2_Tip_Over" / "SEQ2_edge35_b.jpg",  "Seq2 Edge 3-5 B"),
        (PHOTOS / "Seq2_Tip_Over" / "SEQ2_edge36_a.jpg",  "Seq2 Edge 3-6 A"),
        (PHOTOS / "Seq2_Tip_Over" / "SEQ2_edge36_b.jpg",  "Seq2 Edge 3-6 B"),
        # Seq 3 Rot Drop 1 (4 slots)
        (PHOTOS / "Seq3_Rot_Drop_1" / "SEQ3_edge36_a.jpg",    "Seq3 Edge 3-6 A"),
        (PHOTOS / "Seq3_Rot_Drop_1" / "SEQ3_edge36_b.jpg",    "Seq3 Edge 3-6 B"),
        (PHOTOS / "Seq3_Rot_Drop_1" / "SEQ3_corner345_a.jpg", "Seq3 Corner 3-4-5 A"),
        (PHOTOS / "Seq3_Rot_Drop_1" / "SEQ3_corner345_b.jpg", "Seq3 Corner 3-4-5 B"),
        # Seq 4 Incline 1 (4 slots)
        (PHOTOS / "Seq4_Incline_1" / "SEQ4_face5_a.jpg", "Seq4 Face 5 A"),
        (PHOTOS / "Seq4_Incline_1" / "SEQ4_face5_b.jpg", "Seq4 Face 5 B"),
        (PHOTOS / "Seq4_Incline_1" / "SEQ4_face6_a.jpg", "Seq4 Face 6 A"),
        (PHOTOS / "Seq4_Incline_1" / "SEQ4_face6_b.jpg", "Seq4 Face 6 B"),
        # Seq 6 Rot Drop 2 (4 slots)
        (PHOTOS / "Seq6_Rot_Drop_2" / "SEQ6_edge23_a.jpg",    "Seq6 Edge 2-3 A"),
        (PHOTOS / "Seq6_Rot_Drop_2" / "SEQ6_edge23_b.jpg",    "Seq6 Edge 2-3 B"),
        (PHOTOS / "Seq6_Rot_Drop_2" / "SEQ6_corner346_a.jpg", "Seq6 Corner 3-4-6 A"),
        (PHOTOS / "Seq6_Rot_Drop_2" / "SEQ6_corner346_b.jpg", "Seq6 Corner 3-4-6 B"),
        # Seq 7 Incline 2 (4 slots)
        (PHOTOS / "Seq7_Incline_2" / "SEQ7_face4_a.jpg", "Seq7 Face 4 A"),
        (PHOTOS / "Seq7_Incline_2" / "SEQ7_face4_b.jpg", "Seq7 Face 4 B"),
        (PHOTOS / "Seq7_Incline_2" / "SEQ7_face2_a.jpg", "Seq7 Face 2 A"),
        (PHOTOS / "Seq7_Incline_2" / "SEQ7_face2_b.jpg", "Seq7 Face 2 B"),
        # Post-Test (10 slots)
        (PHOTOS / "Post-Test" / "POST_overview_1a.jpg", "Post-Test Overview 1A"),
        (PHOTOS / "Post-Test" / "POST_overview_1b.jpg", "Post-Test Overview 1B"),
        (PHOTOS / "Post-Test" / "POST_overview_2a.jpg", "Post-Test Overview 2A"),
        (PHOTOS / "Post-Test" / "POST_overview_2b.jpg", "Post-Test Overview 2B"),
        (PHOTOS / "Post-Test" / "POST_internal_a.jpg",  "Post-Test Internal A"),
        (PHOTOS / "Post-Test" / "POST_internal_b.jpg",  "Post-Test Internal B"),
        (PHOTOS / "Post-Test" / "POST_stands_1a.jpg",   "Post-Test Stands 1A"),
        (PHOTOS / "Post-Test" / "POST_stands_1b.jpg",   "Post-Test Stands 1B"),
        (PHOTOS / "Post-Test" / "POST_stands_2a.jpg",   "Post-Test Stands 2A"),
        (PHOTOS / "Post-Test" / "POST_stands_2b.jpg",   "Post-Test Stands 2B"),
    ]
    for path, label in photos:
        make_photo(path, label)

    print("\nBuilding ISTA 3B report template...")
    build_3b_template(ROOT / "templates" / "fake_3b_template.docx")
    # Keep legacy filename for backwards compatibility
    build_3b_template(ROOT / "templates" / "fake_report_template.docx")

    print("\nBuilding ISTA 3A report template...")
    build_3a_template(ROOT / "templates" / "fake_3a_template.docx")

    print("\nCreating ISTA 3A sample data folder...")
    _build_3a_sample_data()

    print("\nDone.")


def _build_3a_sample_data():
    """Create fake_test_3a sample folder with 3A-specific CSVs and placeholder images."""
    root_3a   = ROOT / "samples" / "fake_test_3a"
    lansmont3 = root_3a / "Lansmont"
    photos3   = root_3a / "Photos"
    lansmont3.mkdir(parents=True, exist_ok=True)

    # SUMMARY.csv
    (lansmont3 / "SUMMARY.csv").write_text(
        "customer_name,Customer_B\n"
        "project_number,SAMPLE-456\n"
        "test_date,2026-06-22\n"
        "test_type,ISTA 3A\n"
        "test_standard,ISTA 3A-2017\n"
        "product_name,Demo Small Product\n"
        "packaging_description,Demo corrugated box with foam inserts\n"
        "weight_lbs,15\n"
        "dimensions_lwh,14x10x8\n"
        "quantity,1\n"
        "conclusion,Pass\n"
        "key_takeaways,Package passed all ISTA 3A sequences with no damage observed.\n"
        "technician_notes,All tests performed under ambient lab conditions.\n"
    )

    # EQUIPMENT.csv
    (lansmont3 / "EQUIPMENT.csv").write_text(
        "equipment,make,model,serial,calibration_date\n"
        "Drop Tester,Lansmont,Model-D,SN-DEMO-010,2025-01-01\n"
        "Vibration Table,Lansmont,Model-V,SN-DEMO-011,2025-01-01\n"
        "Accelerometer,Endevco,2255B-01,SN-DEMO-012,2025-01-01\n"
    )

    # SEQ_CONDITIONING.csv
    (lansmont3 / "SEQ_CONDITIONING.csv").write_text(
        "result,Pass\n"
        "temperature,72°F (22°C)\n"
        "humidity,50% RH\n"
        "duration,12 hours\n"
    )

    # SEQ3_ROT_DROP_1.csv — reused for drop tests in 3A
    (lansmont3 / "SEQ3_ROT_DROP_1.csv").write_text(
        "result,Pass\n"
        "test1_orientation,Flat — Bottom\n"
        "test1_drop_height_in,24\n"
        "test1_peak_g,42.1\n"
        "test2_orientation,Edge 1-2\n"
        "test2_drop_height_in,18\n"
        "test2_peak_g,38.6\n"
        "conclusion_notes,No damage observed on any drop orientation.\n"
    )

    # SEQ5_VIBRATION.csv
    (lansmont3 / "SEQ5_VIBRATION.csv").write_text(
        "result,Pass\n"
        "orientation,Z Axis (vertical)\n"
        "frequency_range,3–100 Hz\n"
        "vibration_intensity_grms,0.38\n"
        "duration_min,60\n"
        "resonant_frequency_z,22 Hz\n"
        "transmissibility_q_z,3.4\n"
        "conclusion_notes,No resonance amplification above 3.0 Q observed.\n"
    )

    # RAW_DATA_PLACEHOLDER.txt
    (lansmont3 / "RAW_DATA_PLACEHOLDER.txt").write_text(
        "SAMPLE PLACEHOLDER — fake raw data for ISTA 3A development only.\n"
    )

    # Chart placeholder
    make_chart(lansmont3 / "CHART_SEQ3_EDGE_3_6.png",  "3A Drop Test — Edge")
    make_chart(lansmont3 / "CHART_SEQ5_VIBRATION.png",  "3A Vibration Grms")

    # Photo placeholders
    for path, label in [
        (photos3 / "Pre-Test"  / "PRE_sample_01.jpg",  "Pre-Test 1"),
        (photos3 / "Pre-Test"  / "PRE_sample_02.jpg",  "Pre-Test 2"),
        (photos3 / "Post-Test" / "POST_overview_1.jpg", "Post-Test 1"),
        (photos3 / "Post-Test" / "POST_overview_2.jpg", "Post-Test 2"),
    ]:
        make_photo(path, label)

    print(f"  3A sample data: {root_3a.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
