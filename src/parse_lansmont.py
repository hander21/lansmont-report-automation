"""Parses all Lansmont ISTA 3B input files into a structured data dict."""

import logging
from pathlib import Path

import pandas as pd

logger = logging.getLogger(__name__)

# Fields that must be present and non-empty in SUMMARY.csv
REQUIRED_SUMMARY_FIELDS = {
    "customer_name",
    "project_number",
    "test_date",
    "test_type",
    "test_standard",
    "product_name",
    "packaging_description",
    "weight_lbs",
    "dimensions_lwh",
    "quantity",
    "conclusion",
}

OPTIONAL_SUMMARY_FIELDS = {
    "key_takeaways",
    "technician_notes",
    "operator_1",       # "Name | email@transpak.com"
    "operator_2",
    "operator_3",
    "revision",         # "A", "B", "C", …
    "test_method",      # e.g. "ISTA 3B-2021" — replaces the method line on the cover
    "objective",        # Full objective paragraph text (leave blank to keep template default)
}

# Known sequence file stems → prefix used in flattened keys
SEQUENCE_MAP = {
    "SEQ2_TIP_OVER": "seq2",
    "SEQ3_ROT_DROP_1": "seq3",
    "SEQ4_INCLINE_1": "seq4",
    "SEQ5_VIBRATION": "seq5",
    "SEQ6_ROT_DROP_2": "seq6",
    "SEQ7_INCLINE_2": "seq7",
}

EQUIPMENT_COLUMNS = ["equipment", "make", "model", "serial", "calibration_date"]


class ParseError(Exception):
    """Raised when an input file cannot be parsed or is missing required fields."""


# ─── Public entry point ───────────────────────────────────────────────────────

def parse_all(lansmont_folder: Path, summary_path: "Path | None") -> dict:
    """
    Parse the full set of ISTA 3B input files and return a flat dict suitable
    for template substitution and manifest logging.

    Keys from SUMMARY.csv are stored at the top level.
    Equipment rows produce keys  equip_1_equipment … equip_N_calibration_date.
    Sequence fields produce keys seq2_result, seq3_drop_height_inches, etc.

    summary_path may be None — all summary fields will be blank in that case.
    """
    if summary_path is None:
        logger.warning("No summary file — report fields will be blank.")
        parsed: dict = {}
    else:
        parsed = parse_summary(summary_path)

    # Equipment (optional — warn if absent)
    equip_path = lansmont_folder / "EQUIPMENT.csv"
    if equip_path.exists():
        equipment_rows = parse_equipment(equip_path)
        parsed["_equipment_rows"] = equipment_rows          # structured, for manifest
        for i, row in enumerate(equipment_rows, start=1):
            for col in EQUIPMENT_COLUMNS:
                parsed[f"equip_{i}_{col}"] = row.get(col, "")
    else:
        logger.warning("EQUIPMENT.csv not found in %s — equipment table will be blank.", lansmont_folder)
        parsed["_equipment_rows"] = []

    # Sequence CSVs (optional — warn per missing file)
    parsed["_sequences"] = {}
    for stem, prefix in SEQUENCE_MAP.items():
        seq_path = lansmont_folder / f"{stem}.csv"
        if seq_path.exists():
            seq_data = parse_sequence(seq_path, stem)
            parsed["_sequences"][stem] = seq_data
            for k, v in seq_data.items():
                parsed[f"{prefix}_{k}"] = v
        else:
            logger.warning("Sequence file %s.csv not found — %s fields will be blank.", stem, prefix)

    logger.info("All available Lansmont files parsed from %s", lansmont_folder)
    return parsed


def parse_summary(summary_path: Path) -> dict:
    """
    Parse SUMMARY.csv (two-column key,value format).
    Raises ParseError if any required field is missing or blank.
    Never invents, estimates, or defaults any value.
    """
    ext = summary_path.suffix.lower()
    if ext == ".csv":
        raw = _parse_kv_csv(summary_path)
    elif ext in (".xlsx", ".xls"):
        raw = _parse_kv_xlsx(summary_path)
    elif ext == ".txt":
        raw = _parse_kv_txt(summary_path)
    else:
        raise ParseError(
            f"Unsupported summary file format: {ext}. Supported: .csv, .xlsx, .xls, .txt"
        )

    parsed = {k.strip().lower(): str(v).strip() for k, v in raw.items() if k.strip()}

    missing = REQUIRED_SUMMARY_FIELDS - set(parsed.keys())
    if missing:
        raise ParseError(
            f"SUMMARY file is missing required fields: {sorted(missing)}\n"
            f"File: {summary_path}\n"
            "Do NOT guess missing values. Add them to the source file."
        )

    blank = [f for f in REQUIRED_SUMMARY_FIELDS if not parsed.get(f, "").strip()]
    if blank:
        raise ParseError(
            f"These required fields are present but empty: {sorted(blank)}\n"
            f"File: {summary_path}\n"
            "Do NOT fill in blank values automatically. Fix the source file."
        )

    logger.info("SUMMARY parsed successfully from %s", summary_path.name)
    return parsed


def parse_equipment(equip_path: Path) -> list[dict]:
    """
    Parse EQUIPMENT.csv — a tabular CSV with a header row.
    Returns a list of dicts, one per equipment row.
    """
    try:
        df = pd.read_csv(equip_path, dtype=str)
    except Exception as e:
        raise ParseError(f"Failed to read EQUIPMENT.csv at {equip_path}: {e}") from e

    # Normalise column names
    df.columns = [c.strip().lower() for c in df.columns]

    missing_cols = set(EQUIPMENT_COLUMNS) - set(df.columns)
    if missing_cols:
        raise ParseError(
            f"EQUIPMENT.csv is missing required columns: {sorted(missing_cols)}\n"
            f"Expected: {EQUIPMENT_COLUMNS}"
        )

    rows = []
    for _, row in df.iterrows():
        entry = {col: str(row[col]).strip() if pd.notna(row[col]) else "" for col in EQUIPMENT_COLUMNS}
        if any(entry.values()):  # skip fully blank rows
            rows.append(entry)

    if not rows:
        raise ParseError(f"EQUIPMENT.csv has no data rows: {equip_path}")

    logger.info("Equipment parsed: %d instrument(s) from %s", len(rows), equip_path.name)
    return rows


def parse_sequence(seq_path: Path, stem: str) -> dict:
    """
    Parse a per-sequence CSV (two-column key,value format).
    Returns a flat dict of field → value.
    """
    try:
        raw = _parse_kv_csv(seq_path)
    except Exception as e:
        raise ParseError(f"Failed to read sequence file {seq_path}: {e}") from e

    data = {k.strip().lower(): str(v).strip() for k, v in raw.items() if k.strip()}

    if "result" not in data:
        raise ParseError(
            f"Sequence file {seq_path.name} is missing required field 'result'.\n"
            "Every sequence must have a Pass/Fail result. Fix the source file."
        )
    if not data.get("result", "").strip():
        raise ParseError(
            f"Sequence file {seq_path.name} has a blank 'result' field.\n"
            "Do NOT leave the result empty. Fix the source file."
        )

    logger.info("Sequence %s parsed: %d field(s)", stem, len(data))
    return data


# ─── Private CSV/XLSX/TXT readers ────────────────────────────────────────────

def _parse_kv_csv(path: Path) -> dict:
    """Read a two-column key,value CSV (no header row). Values may contain commas."""
    import csv

    try:
        with open(path, newline="", encoding="utf-8-sig") as fh:
            rows = list(csv.reader(fh))
    except Exception as e:
        raise ParseError(f"Failed to read CSV file {path}: {e}") from e

    result = {}
    for cols in rows:
        if not cols:
            continue
        key = cols[0].strip()
        value = ",".join(cols[1:]).strip() if len(cols) > 1 else ""
        if key and key.lower() != "nan":
            result[key] = value
    return result


def _parse_kv_xlsx(path: Path) -> dict:
    """Read a two-column key/value Excel file."""
    try:
        df = pd.read_excel(path, header=None, names=["key", "value"], dtype=str)
    except Exception as e:
        raise ParseError(f"Failed to read XLSX file {path}: {e}") from e

    result = {}
    for _, row in df.iterrows():
        key = str(row["key"]).strip()
        value = str(row["value"]).strip() if pd.notna(row["value"]) else ""
        if key and key.lower() != "nan":
            result[key] = value
    return result


def _parse_kv_txt(path: Path) -> dict:
    """Read a key: value or key=value plain text file."""
    result = {}
    try:
        lines = path.read_text(encoding="utf-8").splitlines()
    except Exception as e:
        raise ParseError(f"Failed to read TXT file {path}: {e}") from e

    for line in lines:
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        for sep in (":", "="):
            if sep in line:
                key, _, value = line.partition(sep)
                result[key.strip()] = value.strip()
                break

    return result
