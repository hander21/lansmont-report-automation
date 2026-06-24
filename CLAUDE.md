# Lansmont Report Automation Rules

This project automates packaging test report drafts from Lansmont-style exported files.

This project may eventually process NDA-covered packaging test data, customer product photos,
customer names, and internal company information.

## Confidentiality Rules (MANDATORY)

Claude must only work with synthetic, fake, or redacted sample files.

Do NOT request:
- Real customer data or customer names
- Real product photos
- Real Lansmont exports or test results
- Confidential report templates
- Internal company information

The generated Python tool must run locally on company-approved machines and process real files
without sending them to any AI model.

The AI is used ONLY to help build and debug the code using fake data. It is NOT used at runtime
to read, summarize, classify, or interpret confidential test files.

## Hard Rules

- Never invent test values
- Never estimate missing values
- Never calculate official test results with AI
- All official report values must come from parsed local source files
- Every value inserted into the report must be logged in manifest.json
- If a required field is missing: stop and show a clear error
- If a required chart is missing: stop and show a clear error
- If required photos are missing: stop and show a warning or error depending on config
- Insert Lansmont-generated charts exactly as exported — do not alter them
- Insert pre-test and post-test photos from local files — do not alter them
- Output reports must be marked DRAFT until reviewed by a human
- Do not delete raw data files
- Preserve original exports and photos
- Prioritize reliability and traceability over flashiness

## Runtime Behavior

- The tool is deterministic: same inputs always produce the same output
- The tool never calls external APIs or sends data over the network
- The tool never uses AI models at runtime
- The tool never emails or distributes reports automatically

## Sample Data Policy

All files under `samples/` are fake/synthetic data only.
- Customer names: Customer_A, Customer_B, etc.
- Product names: Demo_Product, Sample_Package, etc.
- Project numbers: SAMPLE-123, SAMPLE-456, etc.
- Test values: clearly fake numeric samples
- Photos: placeholder images only

## Report Visual Specification (MANDATORY — replicate exactly)

The report template (`templates/fake_report_template.docx`) must always match the
TransPak ISTA 3B report format. When regenerating the template via
`python -m scripts.create_sample_data`, these specs must be preserved exactly.

### Colors
| Element                                  | Hex      | Usage                            |
|------------------------------------------|----------|----------------------------------|
| Header bar + section heading text        | #D31245  | TransPak crimson                 |
| Table section-header rows                | #C00000  | Deep red, white text             |
| Image caption rows                       | #C00000  | Deep red, white bold text        |
| Conclusion header rows                   | #C00000  | Deep red, white text             |
| Equipment table column-header row        | #E7E6E6  | Light gray, black text           |

### Fonts
| Element                   | Font              | Size  | Style         | Color    |
|---------------------------|-------------------|-------|---------------|----------|
| Section headings          | Times New Roman   | 18 pt | Bold          | #D31245  |
| Table section-header text | Calibri           | 12 pt | Bold          | White    |
| Table column-header text  | Calibri           | 12 pt | Bold          | Black    |
| Equipment table col hdrs  | Calibri           | 11 pt | Bold          | Black    |
| Body / data text          | Calibri           | 12 pt | Regular       | Black    |
| Equipment data rows       | Calibri           | 11 pt | Regular       | Black    |
| Footer / page number      | Calibri           | 11 pt | Regular       | Black    |
| Image caption text        | Calibri           | 11 pt | Bold          | White    |

### Header (every page)
- TransPak logo (`templates/transpak_logo.png`), 2.1" wide, left-aligned
- Full-width crimson stripe (#D31245), ~5 pt tall, spanning page edge to page edge
  (achieved via `left_indent = Inches(-1.0)`, `right_indent = Inches(-1.0)` with
   XML `w:shd` paragraph shading)

### Footer (every page)
- "Page | N" right-aligned, Calibri 11pt, using a Word PAGE auto-number field

### Document Margins
- Top: 1.5" (accommodates logo + crimson bar)
- Bottom: 0.75"
- Left / Right: 1.0" each
- Header distance from top edge: 0.25"

### Table Column Widths
| Table             | Column widths (inches)        |
|-------------------|-------------------------------|
| Test Protocol     | 0.7, 1.5, 1.8, 2.5            |
| Equipment         | 2.0, 1.2, 1.0, 1.0, 1.3      |
| Test Sample Desc  | 2.5, 1.0, 2.0, 1.0            |
| Seq 2 (Tip Over)  | 0.8, 3.7, 2.0                 |
| Seq 3/6 (Rot Drop)| 0.8, 2.5, 1.6, 1.6            |
| Seq 4/7 (Incline) | 0.8, 2.0, 1.6, 2.1            |
| Vibration (Seq 5) | 1.8, 1.7, 1.5, 1.5            |
| Photo pairs       | 3.25, 3.25                    |
| Conclusion        | 4.5, 2.0                      |

### Template Token Syntax
- Text fields: `{{ field_name }}` — replaced at runtime from parsed CSVs
- Photo slots:  `[IMG: slot_name]`   — replaced with a copied photo file
- Chart slots:  `[CHART: slot_name]` — replaced with a copied Lansmont chart file

### Chart Insertion
Charts are inserted as images directly from Lansmont-exported PNG files.
No screenshots are needed — drop the exported PNG files into the `Lansmont/`
folder with names like `CHART_SEQ3_EDGE_3_6.png` and the tool inserts them
automatically into the correct slot in the report.

### Regenerating the Template
If the template needs to be rebuilt from scratch:
```bash
python -m scripts.create_sample_data
```
This regenerates both the sample placeholder images and `templates/fake_report_template.docx`
with the exact colors, fonts, header/footer, and layout described above.

## Development Workflow

See README.md for setup and usage instructions.
