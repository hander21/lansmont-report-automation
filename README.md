# Lansmont Report Automation

A local Python CLI tool that automatically generates DRAFT packaging test reports
from Lansmont-exported data, charts, and technician photos.

> **IMPORTANT — NDA / Confidentiality Notice**
> This tool may eventually process NDA-covered test data, customer photos, and
> internal company information. **Real files must never leave the company machine.**
> See [CLAUDE.md](CLAUDE.md) for the full confidentiality rules that govern
> AI-assisted development of this project.

---

## What It Does

1. Reads a completed test folder (Lansmont exports, charts, pre/post photos).
2. Validates all required files are present — stops with a clear error if not.
3. Extracts official test values from the Lansmont summary file.
4. Organizes all files into a clean output folder structure.
5. Inserts charts and photos into a Word report template.
6. Saves the report as a **DRAFT** `.docx` file.
7. Writes a `manifest.json` audit trail recording every source file and value used.

The tool **never** invents, estimates, or calculates test values.
A human engineer **must** review the DRAFT before sending it to any customer.

---

## Setup

### Requirements

- Python 3.11+
- pip

### Install dependencies

```bash
pip install -r requirements.txt
```

### Verify sample data is present

```
samples/fake_test_001/
├── Lansmont/
│   ├── SUMMARY.csv
│   ├── CHART_VIBRATION_Z_AXIS.png
│   ├── CHART_INCLINE_IMPACT.png
│   └── RAW_DATA_PLACEHOLDER.txt
├── Photos/
│   ├── Pre-Test/
│   │   ├── PRE_FRONT_placeholder.jpg
│   │   └── PRE_SIDE_placeholder.jpg
│   └── Post-Test/
│       ├── POST_FRONT_placeholder.jpg
│       └── POST_DAMAGE_placeholder.jpg
└── Notes/
    └── technician_notes.txt
```

All files in `samples/` are **fake/synthetic** — safe for development.

---

## Usage

### Run with default config

```bash
python -m src.main
```

### Run with a specific input folder

```bash
python -m src.main --input samples/fake_test_001
```

### Run with verbose logging

```bash
python -m src.main --log-level DEBUG
```

### Run with a custom config

```bash
python -m src.main --config /path/to/config.yaml
```

---

## Configuration — `config.yaml`

| Key | Description |
|-----|-------------|
| `input_folder` | Path to the test folder to process |
| `output_folder` | Root output folder (e.g. synced OneDrive path) |
| `template_path` | Path to the Word report template |
| `required_files` | Which file types are mandatory (true/false) |
| `photo_settings` | Max display size for photos in the report |
| `chart_settings` | Max display size for charts in the report |
| `report_settings.mark_as_draft` | Always keep reports marked DRAFT |

---

## Output Structure

```
output/Packaging Test Reports/
└── Customer_A/
    └── SAMPLE-123/
        └── 2026-06-22 - VibrationTable/
            ├── Raw Data/          ← original summary + raw files
            ├── Charts/            ← Lansmont chart images
            ├── Photos/
            │   ├── Pre-Test/
            │   └── Post-Test/
            └── Final Report/
                ├── DRAFT_Customer_A_SAMPLE-123_VibrationTable_2026-06-22.docx
                └── manifest.json
```

---

## Input File Naming Conventions

For best results, use these naming conventions in the input folder:

**Summary files** (one required):
- `SUMMARY.csv`
- `TEST_SUMMARY.csv`
- `LANSMONT_SUMMARY.csv`

**Chart images** (from Lansmont export):
- `CHART_VIBRATION_Z_AXIS.png`
- `CHART_INCLINE_IMPACT.png`
- `CHART_SHOCK_RESPONSE.png`

**Pre-test photos**:
- `PRE_FRONT.jpg`
- `PRE_SIDE.jpg`
- `PRE_LABEL.jpg`

**Post-test photos**:
- `POST_FRONT.jpg`
- `POST_DAMAGE.jpg`

If the tool cannot identify a file, it will stop and ask you to rename it.

---

## SUMMARY.csv Format

The summary file uses a two-column `field_name,value` format (no header row):

```csv
customer_name,Customer_A
project_number,SAMPLE-123
test_date,2026-06-22
test_type,Vibration Table
package_description,Demo packaging sample
test_standard,Fake Test Standard 123
peak_g,38.4
duration,120 seconds
axis,Z
result,Pass
technician_notes,Optional free-text notes here
```

Required fields: `customer_name`, `project_number`, `test_date`, `test_type`,
`package_description`, `test_standard`, `peak_g`, `duration`, `axis`, `result`.

---

## manifest.json

Every generated report is accompanied by `manifest.json` in the Final Report folder.
It records:

- Every source file used
- Every value inserted into the report and which field it came from
- Every chart and photo inserted
- Any warnings raised during processing
- Timestamp of generation
- Human review requirement flag

---

## Running Tests

```bash
pytest tests/ -v
```

With coverage:

```bash
pytest tests/ -v --cov=src --cov-report=term-missing
```

---

## Human Review

The tool produces a **DRAFT** report only.

Before sending any report to a customer:
- A qualified test engineer must open the DRAFT file.
- Verify all inserted values match the original Lansmont export.
- Verify all charts and photos are correctly placed.
- Remove the DRAFT watermark and sign off.

The tool will never automatically email or distribute reports.

---

## Future Scope (not yet built)

- Microsoft Graph API / SharePoint integration
- Additional test types: incline impact, drop, shock, compression
- PDF export
- Web interface
- Automatic DRAFT approval workflow

---

## Project Structure

```
lansmont-report-automation/
├── CLAUDE.md                  ← AI dev rules (confidentiality)
├── README.md
├── requirements.txt
├── config.yaml
├── templates/
│   └── fake_report_template.docx
├── samples/
│   └── fake_test_001/         ← synthetic test data only
├── output/                    ← generated reports (gitignored)
├── src/
│   ├── main.py                ← CLI entry point
│   ├── config.py              ← config loader
│   ├── validate_input.py      ← input validation
│   ├── parse_lansmont.py      ← summary file parser
│   ├── organize_files.py      ← file copy/organization
│   ├── generate_report.py     ← Word report generation
│   ├── insert_images.py       ← image sizing/insertion helpers
│   ├── audit_manifest.py      ← manifest.json writer
│   └── utils.py               ← shared utilities
└── tests/
    ├── test_validate_input.py
    ├── test_parse_lansmont.py
    └── test_generate_report.py
```
