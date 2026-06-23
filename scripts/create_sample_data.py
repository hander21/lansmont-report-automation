"""
Creates all synthetic sample images and the fake ISTA 3B report template.
Run from repo root: python -m scripts.create_sample_data
"""

from pathlib import Path
from PIL import Image, ImageDraw, ImageFont
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import shutil

ROOT = Path(__file__).parent.parent
SAMPLES = ROOT / "samples" / "fake_test_001"
LANSMONT = SAMPLES / "Lansmont"
PHOTOS = SAMPLES / "Photos"

# ─── Image helpers ────────────────────────────────────────────────────────────

def make_placeholder_image(path: Path, label: str, color: tuple, size=(800, 600)):
    path.parent.mkdir(parents=True, exist_ok=True)
    img = Image.new("RGB", size, color=color)
    draw = ImageDraw.Draw(img)
    # Draw border
    draw.rectangle([4, 4, size[0]-5, size[1]-5], outline=(0, 0, 0), width=3)
    # Center text
    text = f"[SAMPLE DATA — NOT REAL]\n{label}"
    bbox = draw.textbbox((0, 0), text, font=None)
    tw = bbox[2] - bbox[0]
    th = bbox[3] - bbox[1]
    x = (size[0] - tw) // 2
    y = (size[1] - th) // 2
    draw.text((x, y), text, fill=(30, 30, 30), align="center")
    suffix = path.suffix.lower()
    fmt = "JPEG" if suffix in (".jpg", ".jpeg") else "PNG"
    img.save(str(path), fmt)
    print(f"  Created: {path.relative_to(ROOT)}")


def make_chart(path: Path, label: str):
    make_placeholder_image(path, f"CHART\n{label}", color=(200, 220, 255), size=(1000, 700))


def make_photo(path: Path, label: str):
    make_placeholder_image(path, f"PHOTO\n{label}", color=(220, 240, 220), size=(800, 600))


# ─── Template helpers ─────────────────────────────────────────────────────────

def set_cell_bg(cell, hex_color: str):
    """Set table cell background color."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)


def header_row(table, text: str, cols: int, bg="2C3E50"):
    """Add a full-width header row to a table."""
    row = table.add_row()
    cell = row.cells[0]
    # Merge across all columns
    for i in range(1, cols):
        cell = cell.merge(row.cells[i])
    cell.text = text
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.runs[0]
    run.bold = True
    run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    run.font.size = Pt(10)
    set_cell_bg(cell, bg)


def kv_row(table, label: str, token: str, label_bg="F0F0F0"):
    row = table.add_row()
    lc, vc = row.cells[0], row.cells[1]
    lc.text = label
    lc.paragraphs[0].runs[0].bold = True
    lc.paragraphs[0].runs[0].font.size = Pt(9)
    set_cell_bg(lc, label_bg)
    vc.text = token
    vc.paragraphs[0].runs[0].font.size = Pt(9)


def image_row(table, token: str, cols: int = 1):
    row = table.add_row()
    cell = row.cells[0]
    if cols > 1:
        for i in range(1, cols):
            row.cells[i].text = "" if i < cols else ""
    cell.text = token
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.runs[0]
    run.font.size = Pt(8)
    run.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
    return row


def two_image_row(table, token_a: str, token_b: str):
    row = table.add_row()
    for cell, token in zip(row.cells[:2], [token_a, token_b]):
        cell.text = token
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.runs[0]
        run.font.size = Pt(8)
        run.font.color.rgb = RGBColor(0x88, 0x88, 0x88)


def add_section_break(doc):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(6)


# ─── Build the fake template ──────────────────────────────────────────────────

def build_fake_template(out_path: Path):
    doc = Document()

    # Set narrow margins
    for section in doc.sections:
        section.top_margin = Inches(0.75)
    section.bottom_margin = Inches(0.75)
    section.left_margin = Inches(1.0)
    section.right_margin = Inches(1.0)

    # ── Draft watermark header ──
    p = doc.add_paragraph("** DRAFT — NOT FOR DISTRIBUTION **")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.runs[0]
    run.bold = True
    run.font.color.rgb = RGBColor(0xCC, 0x00, 0x00)
    run.font.size = Pt(12)

    add_section_break(doc)

    # ── TABLE 1: Project Information ──
    tbl = doc.add_table(rows=1, cols=2)
    tbl.style = "Table Grid"
    tbl.columns[0].width = Inches(2.0)
    tbl.columns[1].width = Inches(4.5)
    header_row(tbl, "PROJECT INFORMATION", 2)
    kv_row(tbl, "Customer", "{{ customer_name }}")
    kv_row(tbl, "Project Number", "{{ project_number }}")
    kv_row(tbl, "Test Date", "{{ test_date }}")
    kv_row(tbl, "Test Standard", "{{ test_standard }}")
    kv_row(tbl, "Product Name", "{{ product_name }}")
    kv_row(tbl, "Packaging Description", "{{ packaging_description }}")
    kv_row(tbl, "Weight (lbs)", "{{ weight_lbs }}")
    kv_row(tbl, "Dimensions (L×W×H)", "{{ dimensions_lwh }}")
    kv_row(tbl, "Sample Quantity", "{{ quantity }}")

    add_section_break(doc)

    # ── TABLE 2: Equipment ──
    tbl2 = doc.add_table(rows=1, cols=5)
    tbl2.style = "Table Grid"
    header_row(tbl2, "TEST EQUIPMENT", 5)
    # Column headers row
    hdr = tbl2.add_row()
    for cell, label in zip(hdr.cells, ["Equipment", "Make", "Model", "Serial #", "Cal. Date"]):
        cell.text = label
        cell.paragraphs[0].runs[0].bold = True
        cell.paragraphs[0].runs[0].font.size = Pt(9)
        set_cell_bg(cell, "E8E8E8")
    # Equipment rows (filled at runtime by generator)
    for i in range(1, 7):
        row = tbl2.add_row()
        for j, col in enumerate(["equipment", "make", "model", "serial", "calibration_date"]):
            row.cells[j].text = f"{{{{ equip_{i}_{col} }}}}"
            row.cells[j].paragraphs[0].runs[0].font.size = Pt(9)

    add_section_break(doc)

    # ── SEQ 2: Tip Over ──
    tbl3 = doc.add_table(rows=1, cols=2)
    tbl3.style = "Table Grid"
    tbl3.columns[0].width = Inches(2.5)
    tbl3.columns[1].width = Inches(4.0)
    header_row(tbl3, "SEQUENCE 2 — TIP OVER TEST", 2, bg="1A5276")
    kv_row(tbl3, "Drop Height (in)", "{{ seq2_height_inches }}")
    kv_row(tbl3, "Weight (lbs)", "{{ seq2_weight_lbs }}")
    kv_row(tbl3, "Orientation", "{{ seq2_orientation }}")
    kv_row(tbl3, "Number of Drops", "{{ seq2_num_drops }}")
    kv_row(tbl3, "Result", "{{ seq2_result }}")
    kv_row(tbl3, "Notes", "{{ seq2_notes }}")

    tbl3p = doc.add_table(rows=1, cols=2)
    tbl3p.style = "Table Grid"
    header_row(tbl3p, "SEQUENCE 2 PHOTOS", 2, bg="1A5276")
    two_image_row(tbl3p, "[[ PHOTO:seq2_1 ]]", "[[ PHOTO:seq2_2 ]]")

    add_section_break(doc)

    # ── SEQ 3: Rotational Drop 1 ──
    tbl4 = doc.add_table(rows=1, cols=2)
    tbl4.style = "Table Grid"
    tbl4.columns[0].width = Inches(2.5)
    tbl4.columns[1].width = Inches(4.0)
    header_row(tbl4, "SEQUENCE 3 — ROTATIONAL DROP TEST 1", 2, bg="1A5276")
    kv_row(tbl4, "Drop Height (in)", "{{ seq3_drop_height_inches }}")
    kv_row(tbl4, "Weight (lbs)", "{{ seq3_weight_lbs }}")
    kv_row(tbl4, "Drop Orientations", "{{ seq3_drop_orientations }}")
    kv_row(tbl4, "Number of Drops", "{{ seq3_num_drops }}")
    kv_row(tbl4, "Result", "{{ seq3_result }}")
    kv_row(tbl4, "Notes", "{{ seq3_notes }}")

    tbl4p = doc.add_table(rows=1, cols=2)
    tbl4p.style = "Table Grid"
    header_row(tbl4p, "SEQUENCE 3 PHOTOS", 2, bg="1A5276")
    two_image_row(tbl4p, "[[ PHOTO:seq3_1 ]]", "[[ PHOTO:seq3_2 ]]")

    add_section_break(doc)

    # ── SEQ 4: Incline Impact 1 ──
    tbl5 = doc.add_table(rows=1, cols=2)
    tbl5.style = "Table Grid"
    tbl5.columns[0].width = Inches(2.5)
    tbl5.columns[1].width = Inches(4.0)
    header_row(tbl5, "SEQUENCE 4 — INCLINE IMPACT TEST 1", 2, bg="1A5276")
    kv_row(tbl5, "Incline Angle (deg)", "{{ seq4_incline_angle_degrees }}")
    kv_row(tbl5, "Weight (lbs)", "{{ seq4_weight_lbs }}")
    kv_row(tbl5, "Impact Velocity (fps)", "{{ seq4_impact_velocity_fps }}")
    kv_row(tbl5, "Number of Impacts", "{{ seq4_num_impacts }}")
    kv_row(tbl5, "Result", "{{ seq4_result }}")
    kv_row(tbl5, "Notes", "{{ seq4_notes }}")

    tbl5p = doc.add_table(rows=1, cols=2)
    tbl5p.style = "Table Grid"
    header_row(tbl5p, "SEQUENCE 4 PHOTOS", 2, bg="1A5276")
    two_image_row(tbl5p, "[[ PHOTO:seq4_1 ]]", "[[ PHOTO:seq4_2 ]]")

    add_section_break(doc)

    # ── SEQ 5: Vibration ──
    tbl6 = doc.add_table(rows=1, cols=2)
    tbl6.style = "Table Grid"
    tbl6.columns[0].width = Inches(2.5)
    tbl6.columns[1].width = Inches(4.0)
    header_row(tbl6, "SEQUENCE 5 — VIBRATION TEST", 2, bg="1A5276")
    kv_row(tbl6, "Duration (min)", "{{ seq5_duration_minutes }}")
    kv_row(tbl6, "PSD Profile", "{{ seq5_psd_profile }}")
    kv_row(tbl6, "Peak G", "{{ seq5_peak_g }}")
    kv_row(tbl6, "Frequency Range (Hz)", "{{ seq5_frequency_range_hz }}")
    kv_row(tbl6, "Result", "{{ seq5_result }}")
    kv_row(tbl6, "Notes", "{{ seq5_notes }}")

    add_section_break(doc)

    # ── SEQ 6: Rotational Drop 2 ──
    tbl7 = doc.add_table(rows=1, cols=2)
    tbl7.style = "Table Grid"
    tbl7.columns[0].width = Inches(2.5)
    tbl7.columns[1].width = Inches(4.0)
    header_row(tbl7, "SEQUENCE 6 — ROTATIONAL DROP TEST 2", 2, bg="1A5276")
    kv_row(tbl7, "Drop Height (in)", "{{ seq6_drop_height_inches }}")
    kv_row(tbl7, "Weight (lbs)", "{{ seq6_weight_lbs }}")
    kv_row(tbl7, "Drop Orientations", "{{ seq6_drop_orientations }}")
    kv_row(tbl7, "Number of Drops", "{{ seq6_num_drops }}")
    kv_row(tbl7, "Result", "{{ seq6_result }}")
    kv_row(tbl7, "Notes", "{{ seq6_notes }}")

    tbl7p = doc.add_table(rows=1, cols=2)
    tbl7p.style = "Table Grid"
    header_row(tbl7p, "SEQUENCE 6 PHOTOS", 2, bg="1A5276")
    two_image_row(tbl7p, "[[ PHOTO:seq6_1 ]]", "[[ PHOTO:seq6_2 ]]")

    add_section_break(doc)

    # ── SEQ 7: Incline Impact 2 ──
    tbl8 = doc.add_table(rows=1, cols=2)
    tbl8.style = "Table Grid"
    tbl8.columns[0].width = Inches(2.5)
    tbl8.columns[1].width = Inches(4.0)
    header_row(tbl8, "SEQUENCE 7 — INCLINE IMPACT TEST 2", 2, bg="1A5276")
    kv_row(tbl8, "Incline Angle (deg)", "{{ seq7_incline_angle_degrees }}")
    kv_row(tbl8, "Weight (lbs)", "{{ seq7_weight_lbs }}")
    kv_row(tbl8, "Impact Velocity (fps)", "{{ seq7_impact_velocity_fps }}")
    kv_row(tbl8, "Number of Impacts", "{{ seq7_num_impacts }}")
    kv_row(tbl8, "Result", "{{ seq7_result }}")
    kv_row(tbl8, "Notes", "{{ seq7_notes }}")

    tbl8p = doc.add_table(rows=1, cols=2)
    tbl8p.style = "Table Grid"
    header_row(tbl8p, "SEQUENCE 7 PHOTOS", 2, bg="1A5276")
    two_image_row(tbl8p, "[[ PHOTO:seq7_1 ]]", "[[ PHOTO:seq7_2 ]]")

    add_section_break(doc)

    # ── Pre-Test Photos ──
    tblPre = doc.add_table(rows=1, cols=2)
    tblPre.style = "Table Grid"
    header_row(tblPre, "PRE-TEST PHOTOS", 2, bg="145A32")
    two_image_row(tblPre, "[[ PHOTO:pre_test_1 ]]", "[[ PHOTO:pre_test_2 ]]")

    add_section_break(doc)

    # ── Accelerometer Photos ──
    tblAccel = doc.add_table(rows=1, cols=2)
    tblAccel.style = "Table Grid"
    header_row(tblAccel, "ACCELEROMETER PLACEMENT PHOTOS", 2, bg="145A32")
    two_image_row(tblAccel, "[[ PHOTO:accel_1 ]]", "[[ PHOTO:accel_2 ]]")

    add_section_break(doc)

    # ── Post-Test Photos ──
    tblPost = doc.add_table(rows=1, cols=2)
    tblPost.style = "Table Grid"
    header_row(tblPost, "POST-TEST PHOTOS", 2, bg="145A32")
    two_image_row(tblPost, "[[ PHOTO:post_test_1 ]]", "[[ PHOTO:post_test_2 ]]")

    add_section_break(doc)

    # ── Conclusion ──
    tblConc = doc.add_table(rows=1, cols=2)
    tblConc.style = "Table Grid"
    tblConc.columns[0].width = Inches(2.0)
    tblConc.columns[1].width = Inches(4.5)
    header_row(tblConc, "CONCLUSION", 2, bg="2C3E50")
    kv_row(tblConc, "Key Takeaways", "{{ key_takeaways }}")
    kv_row(tblConc, "Conclusion", "{{ conclusion }}")
    kv_row(tblConc, "Technician Notes", "{{ technician_notes }}")

    add_section_break(doc)

    # ── APPENDIX I: Shock/Impact Charts ──
    p = doc.add_paragraph("APPENDIX I — SHOCK & IMPACT CHARTS")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.runs[0].bold = True
    p.runs[0].font.size = Pt(11)

    chart_sections = [
        ("APPENDIX I — SEQ 2: TIP OVER", "[[ CHART:seq2_tip_over ]]", None),
        ("APPENDIX I — SEQ 3: ROTATIONAL DROP 1 (Charts A & B)",
         "[[ CHART:seq3_rot_drop_1_a ]]", "[[ CHART:seq3_rot_drop_1_b ]]"),
        ("APPENDIX I — SEQ 4: INCLINE IMPACT 1 (Charts A & B)",
         "[[ CHART:seq4_incline_1_a ]]", "[[ CHART:seq4_incline_1_b ]]"),
        ("APPENDIX I — SEQ 6: ROTATIONAL DROP 2 (Charts A & B)",
         "[[ CHART:seq6_rot_drop_2_a ]]", "[[ CHART:seq6_rot_drop_2_b ]]"),
        ("APPENDIX I — SEQ 7: INCLINE IMPACT 2", "[[ CHART:seq7_incline_2 ]]", None),
    ]

    for title, chart_a, chart_b in chart_sections:
        cols = 2 if chart_b else 1
        tblC = doc.add_table(rows=1, cols=cols)
        tblC.style = "Table Grid"
        header_row(tblC, title, cols, bg="4A235A")
        if chart_b:
            row = tblC.add_row()
            row.cells[0].text = chart_a
            row.cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            row.cells[0].paragraphs[0].runs[0].font.size = Pt(8)
            row.cells[0].paragraphs[0].runs[0].font.color.rgb = RGBColor(0x88, 0x88, 0x88)
            row.cells[1].text = chart_b
            row.cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            row.cells[1].paragraphs[0].runs[0].font.size = Pt(8)
            row.cells[1].paragraphs[0].runs[0].font.color.rgb = RGBColor(0x88, 0x88, 0x88)
        else:
            row = tblC.add_row()
            cell = row.cells[0]
            cell.text = chart_a
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            cell.paragraphs[0].runs[0].font.size = Pt(8)
            cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(0x88, 0x88, 0x88)
        add_section_break(doc)

    # ── APPENDIX II: Vibration Chart ──
    tblV = doc.add_table(rows=1, cols=1)
    tblV.style = "Table Grid"
    header_row(tblV, "APPENDIX II — VIBRATION CHART (SEQ 5)", 1, bg="4A235A")
    row = tblV.add_row()
    row.cells[0].text = "[[ CHART:seq5_vibration ]]"
    row.cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    row.cells[0].paragraphs[0].runs[0].font.size = Pt(8)
    row.cells[0].paragraphs[0].runs[0].font.color.rgb = RGBColor(0x88, 0x88, 0x88)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))
    print(f"  Created template: {out_path.relative_to(ROOT)}")


# ─── Main ─────────────────────────────────────────────────────────────────────

def main():
    print("Creating sample placeholder images...")

    # Charts (Lansmont folder)
    charts = {
        "CHART_SEQ2_TIP_OVER.png": "Seq 2 — Tip Over",
        "CHART_SEQ3_ROT_DROP_1_A.png": "Seq 3 — Rot Drop 1 (A)",
        "CHART_SEQ3_ROT_DROP_1_B.png": "Seq 3 — Rot Drop 1 (B)",
        "CHART_SEQ4_INCLINE_1_A.png": "Seq 4 — Incline 1 (A)",
        "CHART_SEQ4_INCLINE_1_B.png": "Seq 4 — Incline 1 (B)",
        "CHART_SEQ5_VIBRATION.png": "Seq 5 — Vibration",
        "CHART_SEQ6_ROT_DROP_2_A.png": "Seq 6 — Rot Drop 2 (A)",
        "CHART_SEQ6_ROT_DROP_2_B.png": "Seq 6 — Rot Drop 2 (B)",
        "CHART_SEQ7_INCLINE_2.png": "Seq 7 — Incline 2",
    }
    for fname, label in charts.items():
        make_chart(LANSMONT / fname, label)

    # Photos
    photo_slots = [
        (PHOTOS / "Accelerometer" / "ACCEL_placement_1.jpg", "Accel Placement 1"),
        (PHOTOS / "Accelerometer" / "ACCEL_placement_2.jpg", "Accel Placement 2"),
        (PHOTOS / "Pre-Test" / "PRE_FRONT_placeholder.jpg", "Pre-Test Front"),
        (PHOTOS / "Pre-Test" / "PRE_SIDE_placeholder.jpg", "Pre-Test Side"),
        (PHOTOS / "Seq2_Tip_Over" / "SEQ2_drop1.jpg", "Seq 2 Drop 1"),
        (PHOTOS / "Seq2_Tip_Over" / "SEQ2_drop2.jpg", "Seq 2 Drop 2"),
        (PHOTOS / "Seq3_Rot_Drop_1" / "SEQ3_drop1.jpg", "Seq 3 Drop 1"),
        (PHOTOS / "Seq3_Rot_Drop_1" / "SEQ3_drop2.jpg", "Seq 3 Drop 2"),
        (PHOTOS / "Seq4_Incline_1" / "SEQ4_impact1.jpg", "Seq 4 Impact 1"),
        (PHOTOS / "Seq4_Incline_1" / "SEQ4_impact2.jpg", "Seq 4 Impact 2"),
        (PHOTOS / "Seq6_Rot_Drop_2" / "SEQ6_drop1.jpg", "Seq 6 Drop 1"),
        (PHOTOS / "Seq6_Rot_Drop_2" / "SEQ6_drop2.jpg", "Seq 6 Drop 2"),
        (PHOTOS / "Seq7_Incline_2" / "SEQ7_impact1.jpg", "Seq 7 Impact 1"),
        (PHOTOS / "Seq7_Incline_2" / "SEQ7_impact2.jpg", "Seq 7 Impact 2"),
        (PHOTOS / "Post-Test" / "POST_FRONT_placeholder.jpg", "Post-Test Front"),
        (PHOTOS / "Post-Test" / "POST_DAMAGE_placeholder.jpg", "Post-Test Damage"),
    ]
    for path, label in photo_slots:
        make_photo(path, label)

    print("\nBuilding fake ISTA 3B report template...")
    build_fake_template(ROOT / "templates" / "fake_report_template.docx")

    print("\nDone. All sample data and template created.")


if __name__ == "__main__":
    main()
