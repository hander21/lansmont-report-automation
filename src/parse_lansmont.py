"""Parses Lansmont summary files (CSV, XLSX, TXT) into a flat dict of test values."""

import logging
from pathlib import Path

import pandas as pd

logger = logging.getLogger(__name__)

# Every field must appear in the source file.  The tool never fills in defaults.
EXPECTED_FIELDS = {
    "customer_name",
    "project_number",
    "test_date",
    "test_type",
    "package_description",
    "test_standard",
    "peak_g",
    "duration",
    "axis",
    "result",
}

OPTIONAL_FIELDS = {"technician_notes", "conclusion"}


class ParseError(Exception):
    """Raised when the summary file cannot be parsed or is missing required fields."""


def parse_summary(summary_path: Path) -> dict:
    """
    Parse a Lansmont summary file and return a dict mapping field names to values.
    Raises ParseError if any required field is missing or blank.
    Never invents, estimates, or defaults any value.
    """
    ext = summary_path.suffix.lower()

    if ext == ".csv":
        raw = _parse_csv(summary_path)
    elif ext in (".xlsx", ".xls"):
        raw = _parse_xlsx(summary_path)
    elif ext == ".txt":
        raw = _parse_txt(summary_path)
    else:
        raise ParseError(
            f"Unsupported summary file format: {ext}. "
            "Supported: .csv, .xlsx, .xls, .txt"
        )

    # Strip whitespace from all keys and values
    parsed = {k.strip().lower(): str(v).strip() for k, v in raw.items()}

    # Verify required fields
    missing = EXPECTED_FIELDS - set(parsed.keys())
    if missing:
        raise ParseError(
            f"Summary file is missing required fields: {sorted(missing)}\n"
            f"File: {summary_path}\n"
            "Do NOT guess missing values. Add them to the source file."
        )

    # Verify no required field is blank
    blank = [f for f in EXPECTED_FIELDS if not parsed.get(f, "").strip()]
    if blank:
        raise ParseError(
            f"These required fields are present but empty: {sorted(blank)}\n"
            f"File: {summary_path}\n"
            "Do NOT fill in blank values automatically. Fix the source file."
        )

    # Validate numeric fields
    _validate_numeric(parsed, "peak_g", summary_path)

    logger.info("Summary parsed successfully from %s", summary_path.name)
    for field in sorted(EXPECTED_FIELDS):
        logger.debug("  %s = %s", field, parsed[field])

    return parsed


def _parse_csv(path: Path) -> dict:
    """Parse a two-column key=value CSV (no header row needed, just key,value pairs)."""
    try:
        df = pd.read_csv(path, header=None, names=["key", "value"], dtype=str)
    except Exception as e:
        raise ParseError(f"Failed to read CSV file {path}: {e}") from e

    if df.shape[1] < 2:
        raise ParseError(
            f"CSV file {path} does not appear to have key,value columns. "
            "Expected two columns: field_name, value."
        )

    result = {}
    for _, row in df.iterrows():
        key = str(row["key"]).strip()
        value = str(row["value"]).strip() if pd.notna(row["value"]) else ""
        if key and key.lower() != "nan":
            result[key] = value
    return result


def _parse_xlsx(path: Path) -> dict:
    """Parse a two-column key/value Excel file."""
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


def _parse_txt(path: Path) -> dict:
    """Parse a key: value or key=value plain text file."""
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


def _validate_numeric(parsed: dict, field: str, path: Path) -> None:
    value = parsed.get(field, "")
    try:
        float(value)
    except ValueError:
        raise ParseError(
            f"Field '{field}' must be a number, got: '{value}'\n"
            f"File: {path}\n"
            "Fix the value in the source file. Do NOT estimate it."
        )
