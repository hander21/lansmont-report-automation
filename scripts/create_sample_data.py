"""
Creates all synthetic sample images and the ISTA 3B report template.
Template structure mirrors the uploaded Ring_ISTA_3B_20260511.pdf.

Image slots:   [IMG: slot_name]    e.g.  [IMG: pre_1]
Chart slots:   [CHART: slot_name]  e.g.  [CHART: seq3_edge_3_6]
Text tokens:   {{ field_name }}    e.g.  {{ customer_name }}

Run from repo root: python -m scripts.create_sample_data
"""

from pathlib import Path
from PIL import Image, ImageDraw
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

ROOT = Path(__file__).parent.parent
SAMPLES = ROOT / "samples" / "fake_test_001"
LANSMONT = SAMPLES / "Lansmont"
PHOTOS = SAMPLES / "Photos"


# ─── Image helpers ────────────────────────────────────────────────────────────

def _make_img(path: Path, label: str, color: tuple, size=(800, 600)):
    path.parent.mkdir(parents=True, exist_ok=True)
    img = Image.new("RGB", size, color=color)
    draw = ImageDraw.Draw(img)
    draw.rectangle([4, 4, size[0]-5, size[1]-5], outline=(0, 0, 0), width=3)
    text = f"[SAMPLE PLACEHOLDER]\n{label}"
    bbox = draw.textbbox((0, 0), text)
    x = (size[0] - (bbox[2]-bbox[0])) // 2
    y = (size[1] - (bbox[3]-bbox[1])) // 2
    draw.text((x, y), text, fill=(30, 30, 30), align="center")
    fmt = "JPEG" if path.suffix.lower() in (".jpg", ".jpeg") else "PNG"
    img.save(str(path), fmt)
    print(f"  {path.relative_to(ROOT)}")


def make_chart(path: Path, label: str):
    _make_img(path, f"CHART — {label}", (200, 215, 255), (1200, 750))


def make_photo(path: Path, label: str):
    _make_img(path, f"PHOTO — {label}", (215, 240, 215), (900, 675))


# ─── Template helpers ─────────────────────────────────────────────────────────

def _set_cell_bg(cell, hex_color: str):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)


def _set_cell_borders(cell):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    for side in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:color"), "000000")
        tcBorders.append(el)
    tcPr.append(tcBorders)


def _cell_text(cell, text: str, bold=False, size=9, color=None, align=None):
    p = cell.paragraphs[0]
    p.clear()
    if align:
        p.alignment = align
    run = p.add_run(text)
    run.bold = bold
    run.font.size = Pt(size)
    if color:
        run.font.color.rgb = RGBColor(*color)


def _page_break(doc):
    p = doc.add_paragraph()
    run = p.add_run()
    run.add_break(__import__("docx.enum.text", fromlist=["WD_BREAK"]).WD_BREAK.PAGE)


def _section_heading(doc, text: str, level=1):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(12 if level == 1 else 10)
    return p


def _body_text(doc, text: str, italic=False):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.font.size = Pt(9)
    run.italic = italic
    return p


def _add_table(doc, rows: int, cols: int, col_widths=None) -> "Table":
    tbl = doc.add_table(rows=rows, cols=cols)
    tbl.style = "Table Grid"
    if col_widths:
        for i, w in enumerate(col_widths):
            for row in tbl.rows:
                row.cells[i].width = Inches(w)
    return tbl


def _img_cell(cell, slot: str):
    """Mark a cell as an image placeholder with named slot."""
    _cell_text(cell, f"[IMG: {slot}]", size=8, color=(150, 150, 150),
               align=WD_ALIGN_PARAGRAPH.CENTER)


def _chart_cell(cell, slot: str):
    """Mark a cell as a chart placeholder with named slot."""
    _cell_text(cell, f"[CHART: {slot}]", size=8, color=(150, 150, 150),
               align=WD_ALIGN_PARAGRAPH.CENTER)


# ─── Build template ───────────────────────────────────────────────────────────

def build_template(out_path: Path):
    doc = Document()

    # Narrow margins to match the report style
    for sec in doc.sections:
        sec.top_margin = Inches(0.75)
        sec.bottom_margin = Inches(0.75)
        sec.left_margin = Inches(1.0)
        sec.right_margin = Inches(1.0)

    # ══════════════════════════════════════════════════
    # PAGE 1 — Cover
    # ══════════════════════════════════════════════════
    p = doc.add_paragraph("DRAFT — NOT FOR DISTRIBUTION")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.runs[0]; r.bold = True; r.font.size = Pt(10)
    r.font.color.rgb = RGBColor(0xCC, 0, 0)

    doc.add_paragraph()

    tbl = _add_table(doc, 1, 2, [1.8, 4.7])
    r0 = tbl.rows[0]
    _cell_text(r0.cells[0], "Test Date:", bold=True)
    _cell_text(r0.cells[1], "{{ test_date }}")

    doc.add_paragraph()
    p = doc.add_paragraph("{{ product_name }}")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.runs[0].bold = True; p.runs[0].font.size = Pt(14)

    p2 = doc.add_paragraph("ISTA 3B Test Method")
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.runs[0].font.size = Pt(11)

    doc.add_paragraph()
    p3 = doc.add_paragraph("TransPak\n20415 Corsair Blvd.\nHayward, CA 94545")
    p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p3.runs[0].font.size = Pt(9)

    _page_break(doc)

    # ══════════════════════════════════════════════════
    # PAGE 2 — Table of Contents (static)
    # ══════════════════════════════════════════════════
    _section_heading(doc, "Table of Contents")
    toc_items = [
        ("Objective", 3), ("Key Takeaways", 3), ("Test Protocol", 4),
        ("Test Equipment", 5), ("Test Sample Description", 6), ("Pretest Preparation", 7),
        ("Test Sequence 1 — Atmospheric Preconditioning", 10),
        ("Test Sequence 2 — Tip / Tip Over", 10),
        ("Test Sequence 3 — Rotational Drop 1", 11),
        ("Test Sequence 4 — Incline Impact 1", 12),
        ("Test Sequence 5 — Vibration", 13),
        ("Test Sequence 6 — Rotational Drop 2", 14),
        ("Test Sequence 7 — Incline Impact 2", 16),
        ("Post-Test Photos", 17),
        ("Appendix I — Shock & Impact Charts", 18),
        ("Appendix II — Vibration Chart", 27),
    ]
    for title, pg in toc_items:
        p = doc.add_paragraph(style="Normal")
        p.add_run(title).font.size = Pt(9)
        p.add_run(f" {'.'*60} {pg}").font.size = Pt(9)

    p = doc.add_paragraph()
    p.add_run("TransPak is an ISTA certified testing facility – Member ID: 10149").font.size = Pt(8)

    _page_break(doc)

    # ══════════════════════════════════════════════════
    # PAGE 3 — Objective + Key Takeaways
    # ══════════════════════════════════════════════════
    _section_heading(doc, "Objective")
    _body_text(doc,
        "To validate the ability of the packaging and product to withstand shipping environment "
        "hazards utilizing a ISTA 3B-2021 General Simulation Performance Test Procedure. Using one "
        "(1) package system, TransPak performed a series of tests that simulate hazards which may be "
        "present in a standard shipping environment.")

    doc.add_paragraph()
    _section_heading(doc, "Key Takeaways")
    _body_text(doc, "{{ key_takeaways }}")

    _page_break(doc)

    # ══════════════════════════════════════════════════
    # PAGE 4 — Test Protocol
    # ══════════════════════════════════════════════════
    _section_heading(doc, "Test Protocol")
    proto = _add_table(doc, 1, 4, [0.7, 1.5, 1.8, 2.5])
    for cell, h in zip(proto.rows[0].cells, ["Sequence No.", "Test Category", "Test Type", "Test Level"]):
        _cell_text(cell, h, bold=True, size=9)
        _set_cell_bg(cell, "D0D0D0")

    rows = [
        ("1", "Atmospheric", "Preconditioning", "Ambient lab conditions\nfor >12 hours"),
        ("2", "Shock", "Tip / Tip Over", "22 Degree Tip Angle"),
        ("3", "Shock", "Rotational Drop", "9\" on Edge 3-6\n9\" on Corner 3-4-5"),
        ("4", "Shock", "Incline Impact", "1 m/sec on Face 5\n1 m/sec on Face 6"),
        ("5", "Vibration", "Random", f"Overall Grms level of {{{{ vibration_grms }}}}"),
        ("6", "Shock", "Rotational Drop", "9\" on Edge 3-4\n9\" on Corner 3-4-6"),
        ("7", "Shock", "Incline Impact", "1 m/sec on Face 2\n1 m/sec on Face 4"),
    ]
    for vals in rows:
        row = proto.add_row()
        for cell, val in zip(row.cells, vals):
            _cell_text(cell, val, size=9)

    _body_text(doc, "* No top load was required as the product is greater than 72\" in height.", italic=True)

    _page_break(doc)

    # ══════════════════════════════════════════════════
    # PAGE 5 — Test Equipment
    # ══════════════════════════════════════════════════
    _section_heading(doc, "Test Equipment")
    equip_tbl = _add_table(doc, 1, 5, [1.8, 1.2, 1.2, 1.1, 1.2])
    for cell, h in zip(equip_tbl.rows[0].cells,
                        ["Equipment", "Make", "Model", "S/N", "Calibration Date"]):
        _cell_text(cell, h, bold=True, size=9)
        _set_cell_bg(cell, "D0D0D0")
    for i in range(1, 9):
        row = equip_tbl.add_row()
        for cell, col in zip(row.cells,
                              ["equipment", "make", "model", "serial", "calibration_date"]):
            _cell_text(cell, f"{{{{ equip_{i}_{col} }}}}", size=9)

    _page_break(doc)

    # ══════════════════════════════════════════════════
    # PAGE 6 — Test Sample Description
    # ══════════════════════════════════════════════════
    _section_heading(doc, "Test Sample Description")
    desc_tbl = _add_table(doc, 2, 4, [2.5, 1.1, 2.2, 0.7])
    for cell, h in zip(desc_tbl.rows[0].cells,
                        ["External Packaging", "Weight (lbs)", "Outside Dimensions (L x W x H)", "Quantity"]):
        _cell_text(cell, h, bold=True, size=9)
        _set_cell_bg(cell, "D0D0D0")
    row = desc_tbl.rows[1]
    _cell_text(row.cells[0], "{{ packaging_description }}", size=9)
    _cell_text(row.cells[1], "{{ weight_lbs }}", size=9)
    _cell_text(row.cells[2], "{{ dimensions_lwh }}", size=9)
    _cell_text(row.cells[3], "{{ quantity }}", size=9)

    doc.add_paragraph()
    img_tbl = _add_table(doc, 1, 1)
    _img_cell(img_tbl.rows[0].cells[0], "pre_1")

    _page_break(doc)

    # ══════════════════════════════════════════════════
    # PAGE 7 — Pretest Preparation (Accelerometer)
    # ══════════════════════════════════════════════════
    _section_heading(doc, "Pretest Preparation")
    _body_text(doc, "Accelerometer")
    accel_tbl = _add_table(doc, 1, 2, [3.25, 3.25])
    _img_cell(accel_tbl.rows[0].cells[0], "accel_1")
    _img_cell(accel_tbl.rows[0].cells[1], "accel_2")
    _body_text(doc, "See above image for notations of X, Y, Z axes.", italic=True)
    _body_text(doc, "Pretest Photos")

    _page_break(doc)

    # ══════════════════════════════════════════════════
    # PAGES 8–9 — Pretest Photos (6 slots)
    # ══════════════════════════════════════════════════
    pre_tbl = _add_table(doc, 2, 2, [3.25, 3.25])
    for i, slot in enumerate(["pre_2", "pre_3", "pre_4", "pre_5"]):
        r, c = divmod(i, 2)
        _img_cell(pre_tbl.rows[r].cells[c], slot)

    _page_break(doc)

    pre_tbl2 = _add_table(doc, 1, 2, [3.25, 3.25])
    _img_cell(pre_tbl2.rows[0].cells[0], "pre_6")
    _img_cell(pre_tbl2.rows[0].cells[1], "pre_7")

    _page_break(doc)

    # ══════════════════════════════════════════════════
    # PAGE 10 — Seq 1 (Atmospheric) + Seq 2 header
    # ══════════════════════════════════════════════════
    _section_heading(doc, "Test Sequence 1 - Atmospheric Preconditioning")
    _body_text(doc,
        "The package system was exposed to ambient laboratory conditions for a minimum "
        "duration of twelve (12) hours.")

    doc.add_paragraph()
    _section_heading(doc, "Test Sequence 2 - Tip / Tip Over")
    _body_text(doc, "Tip / Tip Over")

    seq2_tbl = _add_table(doc, 1, 3, [1.2, 3.0, 1.8])
    for cell, h in zip(seq2_tbl.rows[0].cells, ["Test No.", "Orientation", "Angle (°)"]):
        _cell_text(cell, h, bold=True, size=9)
        _set_cell_bg(cell, "D0D0D0")
    for i in range(1, 5):
        row = seq2_tbl.add_row()
        _cell_text(row.cells[0], str(i), size=9)
        _cell_text(row.cells[1], f"{{{{ seq2_test{i}_orientation }}}}", size=9)
        _cell_text(row.cells[2], f"{{{{ seq2_test{i}_angle }}}}", size=9)

    doc.add_paragraph()
    # 4 orientation pairs (8 photos total: 2 per orientation)
    for idx, (slot_a, slot_b) in enumerate([
        ("seq2_1", "seq2_2"), ("seq2_3", "seq2_4"),
        ("seq2_5", "seq2_6"), ("seq2_7", "seq2_8"),
    ]):
        ph_tbl = _add_table(doc, 2, 2, [3.25, 3.25])
        _img_cell(ph_tbl.rows[0].cells[0], slot_a)
        _img_cell(ph_tbl.rows[0].cells[1], slot_b)
        orient_token = f"{{{{ seq2_test{idx+1}_orientation }}}}"
        _cell_text(ph_tbl.rows[1].cells[0], orient_token, size=8, align=WD_ALIGN_PARAGRAPH.CENTER)
        _cell_text(ph_tbl.rows[1].cells[1], orient_token, size=8, align=WD_ALIGN_PARAGRAPH.CENTER)

    conc2 = _add_table(doc, 2, 2, [1.5, 5.0])
    _cell_text(conc2.rows[0].cells[0], "Conclusion", bold=True)
    _cell_text(conc2.rows[0].cells[1], "Pass / Fail", bold=True)
    _set_cell_bg(conc2.rows[0].cells[0], "D0D0D0")
    _set_cell_bg(conc2.rows[0].cells[1], "D0D0D0")
    _cell_text(conc2.rows[1].cells[0], "{{ seq2_conclusion_notes }}", size=9)
    _cell_text(conc2.rows[1].cells[1], "{{ seq2_result }}", size=9)

    _page_break(doc)

    # ══════════════════════════════════════════════════
    # PAGE 11 — Seq 3: Rotational Drop 1
    # ══════════════════════════════════════════════════
    _section_heading(doc, "Test Sequence 3 - Rotational Drop Sequence 1")
    _body_text(doc, "Rotational Drops")

    seq3_tbl = _add_table(doc, 1, 4, [1.0, 2.2, 1.5, 1.8])
    for cell, h in zip(seq3_tbl.rows[0].cells,
                        ["Test No.", "Orientation", "Drop Height (in.)", "Accelerometer Peak G's"]):
        _cell_text(cell, h, bold=True, size=9)
        _set_cell_bg(cell, "D0D0D0")
    for i in range(1, 3):
        row = seq3_tbl.add_row()
        _cell_text(row.cells[0], str(i), size=9)
        _cell_text(row.cells[1], f"{{{{ seq3_test{i}_orientation }}}}", size=9)
        _cell_text(row.cells[2], f"{{{{ seq3_test{i}_drop_height_in }}}}", size=9)
        _cell_text(row.cells[3], f"{{{{ seq3_test{i}_peak_g }}}}", size=9)

    _body_text(doc, "Note: The opposite impacted edge/corner was supported with a 4.0\" block.", italic=True)
    doc.add_paragraph()

    ph3 = _add_table(doc, 2, 2, [3.25, 3.25])
    _img_cell(ph3.rows[0].cells[0], "seq3_1")
    _img_cell(ph3.rows[0].cells[1], "seq3_2")
    _cell_text(ph3.rows[1].cells[0], "{{ seq3_test1_orientation }}", size=8, align=WD_ALIGN_PARAGRAPH.CENTER)
    _cell_text(ph3.rows[1].cells[1], "{{ seq3_test2_orientation }}", size=8, align=WD_ALIGN_PARAGRAPH.CENTER)

    conc3 = _add_table(doc, 2, 2, [1.5, 5.0])
    _cell_text(conc3.rows[0].cells[0], "Conclusion", bold=True)
    _cell_text(conc3.rows[0].cells[1], "Pass / Fail", bold=True)
    _set_cell_bg(conc3.rows[0].cells[0], "D0D0D0")
    _set_cell_bg(conc3.rows[0].cells[1], "D0D0D0")
    _cell_text(conc3.rows[1].cells[0], "{{ seq3_conclusion_notes }}", size=9)
    _cell_text(conc3.rows[1].cells[1], "{{ seq3_result }}", size=9)

    _page_break(doc)

    # ══════════════════════════════════════════════════
    # PAGE 12 — Seq 4: Incline Impact 1
    # ══════════════════════════════════════════════════
    _section_heading(doc, "Test Sequence 4 - Incline Impact Sequence 1")
    _body_text(doc, "Incline Impact")

    seq4_tbl = _add_table(doc, 1, 4, [1.0, 1.8, 1.5, 2.2])
    for cell, h in zip(seq4_tbl.rows[0].cells,
                        ["Test No.", "Face", "Velocity", "Accelerometer Peak G's"]):
        _cell_text(cell, h, bold=True, size=9)
        _set_cell_bg(cell, "D0D0D0")
    for i in range(1, 3):
        row = seq4_tbl.add_row()
        _cell_text(row.cells[0], str(i), size=9)
        _cell_text(row.cells[1], f"{{{{ seq4_test{i}_face }}}}", size=9)
        _cell_text(row.cells[2], f"{{{{ seq4_test{i}_velocity }}}}", size=9)
        _cell_text(row.cells[3], f"{{{{ seq4_test{i}_peak_g }}}}", size=9)

    doc.add_paragraph()
    ph4 = _add_table(doc, 2, 2, [3.25, 3.25])
    _img_cell(ph4.rows[0].cells[0], "seq4_1")
    _img_cell(ph4.rows[0].cells[1], "seq4_2")
    _cell_text(ph4.rows[1].cells[0], "{{ seq4_test1_face }}", size=8, align=WD_ALIGN_PARAGRAPH.CENTER)
    _cell_text(ph4.rows[1].cells[1], "{{ seq4_test2_face }}", size=8, align=WD_ALIGN_PARAGRAPH.CENTER)

    conc4 = _add_table(doc, 2, 2, [1.5, 5.0])
    _cell_text(conc4.rows[0].cells[0], "Conclusion", bold=True)
    _cell_text(conc4.rows[0].cells[1], "Pass / Fail", bold=True)
    _set_cell_bg(conc4.rows[0].cells[0], "D0D0D0")
    _set_cell_bg(conc4.rows[0].cells[1], "D0D0D0")
    _cell_text(conc4.rows[1].cells[0], "{{ seq4_conclusion_notes }}", size=9)
    _cell_text(conc4.rows[1].cells[1], "{{ seq4_result }}", size=9)

    _page_break(doc)

    # ══════════════════════════════════════════════════
    # PAGES 13–14 — Seq 5: Vibration
    # ══════════════════════════════════════════════════
    _section_heading(doc, "Test Sequence 5 - Vibration")
    _body_text(doc, "Random Vibration")

    vib_tbl = _add_table(doc, 1, 4, [1.5, 1.5, 1.8, 1.7])
    for cell, h in zip(vib_tbl.rows[0].cells,
                        ["Orientation", "Frequency Range", "Vibration Intensity (Grms)", "Duration (min)"]):
        _cell_text(cell, h, bold=True, size=9)
        _set_cell_bg(cell, "D0D0D0")
    vrow = vib_tbl.add_row()
    _cell_text(vrow.cells[0], "{{ seq5_orientation }}", size=9)
    _cell_text(vrow.cells[1], "{{ seq5_frequency_range }}", size=9)
    _cell_text(vrow.cells[2], "{{ seq5_vibration_intensity_grms }}", size=9)
    _cell_text(vrow.cells[3], "{{ seq5_duration_min }}", size=9)

    _body_text(doc, "Note: No top load was required as the rack system is greater than 72\" in height.", italic=True)
    doc.add_paragraph()

    # PSD spectrum table (fixed values from ISTA 3B standard)
    _body_text(doc, "Random Vibration Spectrum")
    psd_tbl = _add_table(doc, 1, 2, [2.5, 2.5])
    _cell_text(psd_tbl.rows[0].cells[0], "Frequency (Hz)", bold=True, size=9)
    _cell_text(psd_tbl.rows[0].cells[1], "PSD (g² / Hz)", bold=True, size=9)
    _set_cell_bg(psd_tbl.rows[0].cells[0], "D0D0D0")
    _set_cell_bg(psd_tbl.rows[0].cells[1], "D0D0D0")
    psd_vals = [
        ("1", "0.00072"), ("3", "0.018"), ("4", "0.018"), ("6", "0.00072"),
        ("12", "0.00072"), ("16", "0.0036"), ("25", "0.0036"), ("30", "0.00072"),
        ("40", "0.0036"), ("80", "0.0036"), ("100", "0.00036"), ("200", "0.000018"),
    ]
    for freq, psd in psd_vals:
        r = psd_tbl.add_row()
        _cell_text(r.cells[0], freq, size=9)
        _cell_text(r.cells[1], psd, size=9)

    _page_break(doc)

    _body_text(doc, "Accelerometer #1 Results")
    res_tbl = _add_table(doc, 2, 2, [3.25, 3.25])
    _cell_text(res_tbl.rows[0].cells[0], "Resonant Frequency Z Axis", bold=True, size=9)
    _cell_text(res_tbl.rows[0].cells[1], "Transmissibility (Q) Z Axis", bold=True, size=9)
    _set_cell_bg(res_tbl.rows[0].cells[0], "D0D0D0")
    _set_cell_bg(res_tbl.rows[0].cells[1], "D0D0D0")
    _cell_text(res_tbl.rows[1].cells[0], "{{ seq5_resonant_frequency_z }}", size=9)
    _cell_text(res_tbl.rows[1].cells[1], "{{ seq5_transmissibility_q_z }}", size=9)

    _body_text(doc, "Accelerometer #1 Transmissibility Report")
    vib_chart_tbl = _add_table(doc, 1, 1)
    _chart_cell(vib_chart_tbl.rows[0].cells[0], "seq5_vibration")

    doc.add_paragraph()
    conc5 = _add_table(doc, 2, 2, [1.5, 5.0])
    _cell_text(conc5.rows[0].cells[0], "Conclusion", bold=True)
    _cell_text(conc5.rows[0].cells[1], "Pass / Fail", bold=True)
    _set_cell_bg(conc5.rows[0].cells[0], "D0D0D0")
    _set_cell_bg(conc5.rows[0].cells[1], "D0D0D0")
    _cell_text(conc5.rows[1].cells[0], "{{ seq5_conclusion_notes }}", size=9)
    _cell_text(conc5.rows[1].cells[1], "{{ seq5_result }}", size=9)

    _page_break(doc)

    # ══════════════════════════════════════════════════
    # PAGES 14–15 — Seq 6: Rotational Drop 2
    # ══════════════════════════════════════════════════
    _section_heading(doc, "Test Sequence 6 - Rotational Drop Sequence 2")
    _body_text(doc, "Rotational Drops")

    seq6_tbl = _add_table(doc, 1, 4, [1.0, 2.2, 1.5, 1.8])
    for cell, h in zip(seq6_tbl.rows[0].cells,
                        ["Test No.", "Orientation", "Drop Height (in.)", "Accelerometer #2 Peak G's"]):
        _cell_text(cell, h, bold=True, size=9)
        _set_cell_bg(cell, "D0D0D0")
    for i in range(1, 3):
        row = seq6_tbl.add_row()
        _cell_text(row.cells[0], str(i), size=9)
        _cell_text(row.cells[1], f"{{{{ seq6_test{i}_orientation }}}}", size=9)
        _cell_text(row.cells[2], f"{{{{ seq6_test{i}_drop_height_in }}}}", size=9)
        _cell_text(row.cells[3], f"{{{{ seq6_test{i}_peak_g }}}}", size=9)

    _body_text(doc, "Note: The opposite impacted edge/corner was supported with a 4.0\" block.", italic=True)
    doc.add_paragraph()

    ph6 = _add_table(doc, 2, 2, [3.25, 3.25])
    _img_cell(ph6.rows[0].cells[0], "seq6_1")
    _img_cell(ph6.rows[0].cells[1], "seq6_2")
    _cell_text(ph6.rows[1].cells[0], "{{ seq6_test1_orientation }}", size=8, align=WD_ALIGN_PARAGRAPH.CENTER)
    _cell_text(ph6.rows[1].cells[1], "{{ seq6_test2_orientation }}", size=8, align=WD_ALIGN_PARAGRAPH.CENTER)

    _page_break(doc)

    conc6 = _add_table(doc, 2, 2, [1.5, 5.0])
    _cell_text(conc6.rows[0].cells[0], "Conclusion", bold=True)
    _cell_text(conc6.rows[0].cells[1], "Pass / Fail", bold=True)
    _set_cell_bg(conc6.rows[0].cells[0], "D0D0D0")
    _set_cell_bg(conc6.rows[0].cells[1], "D0D0D0")
    _cell_text(conc6.rows[1].cells[0], "{{ seq6_conclusion_notes }}", size=9)
    _cell_text(conc6.rows[1].cells[1], "{{ seq6_result }}", size=9)

    _page_break(doc)

    # ══════════════════════════════════════════════════
    # PAGE 16 — Seq 7: Incline Impact 2
    # ══════════════════════════════════════════════════
    _section_heading(doc, "Test Sequence 7 - Incline Impact Sequence 2")
    _body_text(doc, "Incline Impact")

    seq7_tbl = _add_table(doc, 1, 4, [1.0, 1.8, 1.5, 2.2])
    for cell, h in zip(seq7_tbl.rows[0].cells,
                        ["Test No.", "Face", "Velocity", "Accelerometer Peak G's"]):
        _cell_text(cell, h, bold=True, size=9)
        _set_cell_bg(cell, "D0D0D0")
    for i in range(1, 3):
        row = seq7_tbl.add_row()
        _cell_text(row.cells[0], str(i), size=9)
        _cell_text(row.cells[1], f"{{{{ seq7_test{i}_face }}}}", size=9)
        _cell_text(row.cells[2], f"{{{{ seq7_test{i}_velocity }}}}", size=9)
        _cell_text(row.cells[3], f"{{{{ seq7_test{i}_peak_g }}}}", size=9)

    doc.add_paragraph()
    ph7 = _add_table(doc, 2, 2, [3.25, 3.25])
    _img_cell(ph7.rows[0].cells[0], "seq7_1")
    _img_cell(ph7.rows[0].cells[1], "seq7_2")
    _cell_text(ph7.rows[1].cells[0], "{{ seq7_test2_face }}", size=8, align=WD_ALIGN_PARAGRAPH.CENTER)
    _cell_text(ph7.rows[1].cells[1], "{{ seq7_test1_face }}", size=8, align=WD_ALIGN_PARAGRAPH.CENTER)

    conc7 = _add_table(doc, 2, 2, [1.5, 5.0])
    _cell_text(conc7.rows[0].cells[0], "Conclusion", bold=True)
    _cell_text(conc7.rows[0].cells[1], "Pass / Fail", bold=True)
    _set_cell_bg(conc7.rows[0].cells[0], "D0D0D0")
    _set_cell_bg(conc7.rows[0].cells[1], "D0D0D0")
    _cell_text(conc7.rows[1].cells[0], "{{ seq7_conclusion_notes }}", size=9)
    _cell_text(conc7.rows[1].cells[1], "{{ seq7_result }}", size=9)

    _page_break(doc)

    # ══════════════════════════════════════════════════
    # PAGE 17 — Post-Test Photos (5 pairs = 10 photos)
    # ══════════════════════════════════════════════════
    _section_heading(doc, "Post-Test Photos")
    labels = ["Overview", "Overview", "Internal View", "Stands", "Stands"]
    for pair_idx, label in enumerate(labels):
        slot_a = f"post_{pair_idx*2+1}"
        slot_b = f"post_{pair_idx*2+2}"
        post_tbl = _add_table(doc, 2, 2, [3.25, 3.25])
        _img_cell(post_tbl.rows[0].cells[0], slot_a)
        _img_cell(post_tbl.rows[0].cells[1], slot_b)
        _cell_text(post_tbl.rows[1].cells[0], label, size=8, align=WD_ALIGN_PARAGRAPH.CENTER)
        _cell_text(post_tbl.rows[1].cells[1], label, size=8, align=WD_ALIGN_PARAGRAPH.CENTER)
        doc.add_paragraph()

    _page_break(doc)

    # ══════════════════════════════════════════════════
    # PAGE 18 — Appendix I header
    # ══════════════════════════════════════════════════
    p = doc.add_paragraph("Appendix I")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.runs[0].bold = True; p.runs[0].font.size = Pt(14)

    _page_break(doc)

    # ══════════════════════════════════════════════════
    # PAGES 19–26 — Appendix I charts (8 charts)
    # ══════════════════════════════════════════════════
    appendix_i_charts = [
        ("Rotational Drop Edge 3-6",       "seq3_edge_3_6"),
        ("Rotational Drop Corner 3-4-5",   "seq3_corner_3_4_5"),
        ("Impact Face 5",                  "seq4_face_5"),
        ("Impact Face 6",                  "seq4_face_6"),
        ("Rotational Edge Drop 2-3",       "seq6_edge_2_3"),
        ("Corner Drop 3-4-6",              "seq6_corner_3_4_6"),
        ("Impact Face 4",                  "seq7_face_4"),
        ("Impact Face 2",                  "seq7_face_2"),
    ]
    for label, slot in appendix_i_charts:
        _body_text(doc, label)
        ct = _add_table(doc, 1, 1)
        _chart_cell(ct.rows[0].cells[0], slot)
        _page_break(doc)

    # ══════════════════════════════════════════════════
    # PAGE 27 — Appendix II header
    # ══════════════════════════════════════════════════
    p = doc.add_paragraph("Appendix II")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.runs[0].bold = True; p.runs[0].font.size = Pt(14)

    _page_break(doc)

    # ══════════════════════════════════════════════════
    # PAGE 28 — Vibration chart + End of Report
    # ══════════════════════════════════════════════════
    _body_text(doc, "Vibration Grms {{ seq5_vibration_intensity_grms }} | {{ seq5_duration_min }} minutes")
    vib2 = _add_table(doc, 1, 1)
    _chart_cell(vib2.rows[0].cells[0], "seq5_vibration")

    doc.add_paragraph()
    p = doc.add_paragraph("End of Report")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.runs[0].bold = True; p.runs[0].font.size = Pt(10)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))
    print(f"  Template: {out_path.relative_to(ROOT)}")


# ─── Main ─────────────────────────────────────────────────────────────────────

def main():
    print("Creating chart placeholder images...")
    charts = {
        "CHART_SEQ3_EDGE_3_6.png":      "Seq 3 — Rotational Drop Edge 3-6",
        "CHART_SEQ3_CORNER_3_4_5.png":  "Seq 3 — Rotational Drop Corner 3-4-5",
        "CHART_SEQ4_FACE_5.png":        "Seq 4 — Impact Face 5",
        "CHART_SEQ4_FACE_6.png":        "Seq 4 — Impact Face 6",
        "CHART_SEQ5_VIBRATION.png":     "Seq 5 — Vibration",
        "CHART_SEQ6_EDGE_2_3.png":      "Seq 6 — Rotational Edge Drop 2-3",
        "CHART_SEQ6_CORNER_3_4_6.png":  "Seq 6 — Corner Drop 3-4-6",
        "CHART_SEQ7_FACE_4.png":        "Seq 7 — Impact Face 4",
        "CHART_SEQ7_FACE_2.png":        "Seq 7 — Impact Face 2",
    }
    for fname, label in charts.items():
        make_chart(LANSMONT / fname, label)

    print("\nCreating photo placeholder images...")
    photos = [
        # Accelerometer (page 7)
        (PHOTOS / "Accelerometer" / "ACCEL_placement_1.jpg", "Accelerometer 1"),
        (PHOTOS / "Accelerometer" / "ACCEL_placement_2.jpg", "Accelerometer 2"),
        # Pre-Test: 7 slots (1 for sample description, 6 for pretest pages)
        (PHOTOS / "Pre-Test" / "PRE_sample_desc.jpg",  "Pre-Test — Sample Description"),
        (PHOTOS / "Pre-Test" / "PRE_pretest_02.jpg",   "Pre-Test Photo 2"),
        (PHOTOS / "Pre-Test" / "PRE_pretest_03.jpg",   "Pre-Test Photo 3"),
        (PHOTOS / "Pre-Test" / "PRE_pretest_04.jpg",   "Pre-Test Photo 4"),
        (PHOTOS / "Pre-Test" / "PRE_pretest_05.jpg",   "Pre-Test Photo 5"),
        (PHOTOS / "Pre-Test" / "PRE_pretest_06.jpg",   "Pre-Test Photo 6"),
        (PHOTOS / "Pre-Test" / "PRE_pretest_07.jpg",   "Pre-Test Photo 7"),
        # Seq 2 Tip Over: 8 slots (2 per orientation × 4 orientations)
        (PHOTOS / "Seq2_Tip_Over" / "SEQ2_edge23_a.jpg",  "Seq2 Edge 2-3 A"),
        (PHOTOS / "Seq2_Tip_Over" / "SEQ2_edge23_b.jpg",  "Seq2 Edge 2-3 B"),
        (PHOTOS / "Seq2_Tip_Over" / "SEQ2_edge34_a.jpg",  "Seq2 Edge 3-4 A"),
        (PHOTOS / "Seq2_Tip_Over" / "SEQ2_edge34_b.jpg",  "Seq2 Edge 3-4 B"),
        (PHOTOS / "Seq2_Tip_Over" / "SEQ2_edge35_a.jpg",  "Seq2 Edge 3-5 A"),
        (PHOTOS / "Seq2_Tip_Over" / "SEQ2_edge35_b.jpg",  "Seq2 Edge 3-5 B"),
        (PHOTOS / "Seq2_Tip_Over" / "SEQ2_edge36_a.jpg",  "Seq2 Edge 3-6 A"),
        (PHOTOS / "Seq2_Tip_Over" / "SEQ2_edge36_b.jpg",  "Seq2 Edge 3-6 B"),
        # Seq 3 Rot Drop 1: 4 slots
        (PHOTOS / "Seq3_Rot_Drop_1" / "SEQ3_edge36_a.jpg",   "Seq3 Edge 3-6 A"),
        (PHOTOS / "Seq3_Rot_Drop_1" / "SEQ3_edge36_b.jpg",   "Seq3 Edge 3-6 B"),
        (PHOTOS / "Seq3_Rot_Drop_1" / "SEQ3_corner345_a.jpg","Seq3 Corner 3-4-5 A"),
        (PHOTOS / "Seq3_Rot_Drop_1" / "SEQ3_corner345_b.jpg","Seq3 Corner 3-4-5 B"),
        # Seq 4 Incline 1: 4 slots
        (PHOTOS / "Seq4_Incline_1" / "SEQ4_face5_a.jpg",  "Seq4 Face 5 A"),
        (PHOTOS / "Seq4_Incline_1" / "SEQ4_face5_b.jpg",  "Seq4 Face 5 B"),
        (PHOTOS / "Seq4_Incline_1" / "SEQ4_face6_a.jpg",  "Seq4 Face 6 A"),
        (PHOTOS / "Seq4_Incline_1" / "SEQ4_face6_b.jpg",  "Seq4 Face 6 B"),
        # Seq 6 Rot Drop 2: 4 slots
        (PHOTOS / "Seq6_Rot_Drop_2" / "SEQ6_edge23_a.jpg",   "Seq6 Edge 2-3 A"),
        (PHOTOS / "Seq6_Rot_Drop_2" / "SEQ6_edge23_b.jpg",   "Seq6 Edge 2-3 B"),
        (PHOTOS / "Seq6_Rot_Drop_2" / "SEQ6_corner346_a.jpg","Seq6 Corner 3-4-6 A"),
        (PHOTOS / "Seq6_Rot_Drop_2" / "SEQ6_corner346_b.jpg","Seq6 Corner 3-4-6 B"),
        # Seq 7 Incline 2: 4 slots
        (PHOTOS / "Seq7_Incline_2" / "SEQ7_face4_a.jpg",  "Seq7 Face 4 A"),
        (PHOTOS / "Seq7_Incline_2" / "SEQ7_face4_b.jpg",  "Seq7 Face 4 B"),
        (PHOTOS / "Seq7_Incline_2" / "SEQ7_face2_a.jpg",  "Seq7 Face 2 A"),
        (PHOTOS / "Seq7_Incline_2" / "SEQ7_face2_b.jpg",  "Seq7 Face 2 B"),
        # Post-Test: 10 slots
        (PHOTOS / "Post-Test" / "POST_overview_1a.jpg",   "Post-Test Overview 1A"),
        (PHOTOS / "Post-Test" / "POST_overview_1b.jpg",   "Post-Test Overview 1B"),
        (PHOTOS / "Post-Test" / "POST_overview_2a.jpg",   "Post-Test Overview 2A"),
        (PHOTOS / "Post-Test" / "POST_overview_2b.jpg",   "Post-Test Overview 2B"),
        (PHOTOS / "Post-Test" / "POST_internal_a.jpg",    "Post-Test Internal A"),
        (PHOTOS / "Post-Test" / "POST_internal_b.jpg",    "Post-Test Internal B"),
        (PHOTOS / "Post-Test" / "POST_stands_1a.jpg",     "Post-Test Stands 1A"),
        (PHOTOS / "Post-Test" / "POST_stands_1b.jpg",     "Post-Test Stands 1B"),
        (PHOTOS / "Post-Test" / "POST_stands_2a.jpg",     "Post-Test Stands 2A"),
        (PHOTOS / "Post-Test" / "POST_stands_2b.jpg",     "Post-Test Stands 2B"),
    ]
    for path, label in photos:
        make_photo(path, label)

    print("\nBuilding ISTA 3B report template...")
    build_template(ROOT / "templates" / "fake_report_template.docx")

    print("\nDone.")


if __name__ == "__main__":
    main()
