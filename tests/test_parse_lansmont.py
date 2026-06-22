"""Tests for src/parse_lansmont.py using fake CSV and edge cases."""

import pytest
from pathlib import Path

from src.parse_lansmont import parse_summary, ParseError, EXPECTED_FIELDS

FAKE_SUMMARY = Path("samples/fake_test_001/Lansmont/SUMMARY.csv")


class TestParseSummaryCSV:
    def test_parses_fake_summary_successfully(self):
        result = parse_summary(FAKE_SUMMARY)
        assert result["customer_name"] == "Customer_A"
        assert result["project_number"] == "SAMPLE-123"
        assert result["test_date"] == "2026-06-22"
        assert result["test_type"] == "Vibration Table"
        assert result["peak_g"] == "38.4"
        assert result["result"] == "Pass"

    def test_all_required_fields_present(self):
        result = parse_summary(FAKE_SUMMARY)
        for field in EXPECTED_FIELDS:
            assert field in result, f"Missing required field: {field}"
            assert result[field].strip() != "", f"Field is blank: {field}"

    def test_peak_g_is_numeric_string(self):
        result = parse_summary(FAKE_SUMMARY)
        float(result["peak_g"])  # should not raise

    def test_fails_on_missing_required_field(self, tmp_path):
        csv = tmp_path / "SUMMARY.csv"
        # Write all fields except peak_g
        lines = [f"{f},sample_value" for f in EXPECTED_FIELDS if f != "peak_g"]
        csv.write_text("\n".join(lines))
        with pytest.raises(ParseError, match="peak_g"):
            parse_summary(csv)

    def test_fails_on_blank_required_field(self, tmp_path):
        csv = tmp_path / "SUMMARY.csv"
        lines = [f"{f},sample_value" for f in EXPECTED_FIELDS]
        lines.append("result,")  # blank result
        # Overwrite result line
        content = "\n".join(f"{f},{'' if f == 'result' else 'val'}" for f in EXPECTED_FIELDS)
        csv.write_text(content)
        with pytest.raises(ParseError, match="result"):
            parse_summary(csv)

    def test_fails_on_non_numeric_peak_g(self, tmp_path):
        csv = tmp_path / "SUMMARY.csv"
        content = "\n".join(f"{f},val" for f in EXPECTED_FIELDS)
        content = content.replace("peak_g,val", "peak_g,NOT_A_NUMBER")
        csv.write_text(content)
        with pytest.raises(ParseError, match="peak_g"):
            parse_summary(csv)

    def test_fails_on_unsupported_extension(self, tmp_path):
        f = tmp_path / "SUMMARY.json"
        f.write_text("{}")
        with pytest.raises(ParseError, match="Unsupported"):
            parse_summary(f)

    def test_keys_are_lowercased(self):
        result = parse_summary(FAKE_SUMMARY)
        for key in result:
            assert key == key.lower(), f"Key is not lowercase: {key}"

    def test_values_are_stripped(self):
        result = parse_summary(FAKE_SUMMARY)
        for key, val in result.items():
            assert val == val.strip(), f"Value not stripped for key: {key}"


class TestParseSummaryTXT:
    def test_parses_txt_with_colon_separator(self, tmp_path):
        txt = tmp_path / "SUMMARY.txt"
        fields = {f: "val" for f in EXPECTED_FIELDS}
        fields["peak_g"] = "25.0"
        content = "\n".join(f"{k}: {v}" for k, v in fields.items())
        txt.write_text(content)
        result = parse_summary(txt)
        assert result["peak_g"] == "25.0"
        assert result["customer_name"] == "val"

    def test_parses_txt_with_equals_separator(self, tmp_path):
        txt = tmp_path / "SUMMARY.txt"
        fields = {f: "val" for f in EXPECTED_FIELDS}
        fields["peak_g"] = "10.5"
        content = "\n".join(f"{k}={v}" for k, v in fields.items())
        txt.write_text(content)
        result = parse_summary(txt)
        assert result["peak_g"] == "10.5"

    def test_ignores_comment_lines(self, tmp_path):
        txt = tmp_path / "SUMMARY.txt"
        fields = {f: "val" for f in EXPECTED_FIELDS}
        fields["peak_g"] = "7.7"
        lines = [f"# This is a comment"] + [f"{k}: {v}" for k, v in fields.items()]
        txt.write_text("\n".join(lines))
        result = parse_summary(txt)
        assert result["peak_g"] == "7.7"
