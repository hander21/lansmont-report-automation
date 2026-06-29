"""Tests for src/parse_lansmont.py using fake ISTA 3B CSV files."""

import pytest
from pathlib import Path

from src.parse_lansmont import (
    parse_summary, parse_equipment, parse_sequence, parse_all,
    ParseError, REQUIRED_SUMMARY_FIELDS, EQUIPMENT_COLUMNS,
)

LANSMONT_DIR = Path("samples/fake_test_001/Lansmont")
FAKE_SUMMARY = LANSMONT_DIR / "SUMMARY.csv"
FAKE_EQUIPMENT = LANSMONT_DIR / "EQUIPMENT.csv"
FAKE_SEQ2 = LANSMONT_DIR / "SEQ2_TIP_OVER.csv"


class TestParseSummary:
    def test_parses_fake_summary_successfully(self):
        result = parse_summary(FAKE_SUMMARY)
        assert result["customer_name"] == "Customer_A"
        assert result["project_number"] == "SAMPLE-123"
        assert result["test_date"] == "June 22 2026"
        assert result["test_type"] == "ISTA 3B"
        assert result["product_name"] == "Demo Product Sample"
        assert result["weight_lbs"] == "185 lbs."

    def test_all_required_fields_present(self):
        result = parse_summary(FAKE_SUMMARY)
        for field in REQUIRED_SUMMARY_FIELDS:
            assert field in result, f"Missing required field: {field}"
            assert result[field].strip() != "", f"Field is blank: {field}"

    def test_keys_are_lowercased(self):
        result = parse_summary(FAKE_SUMMARY)
        for key in result:
            assert key == key.lower(), f"Key is not lowercase: {key}"

    def test_values_are_stripped(self):
        result = parse_summary(FAKE_SUMMARY)
        for key, val in result.items():
            assert val == val.strip(), f"Value not stripped for: {key}"

    def test_fails_on_missing_required_field(self, tmp_path):
        csv = tmp_path / "SUMMARY.csv"
        lines = [f"{f},sample_value" for f in REQUIRED_SUMMARY_FIELDS if f != "conclusion"]
        csv.write_text("\n".join(lines))
        with pytest.raises(ParseError, match="conclusion"):
            parse_summary(csv)

    def test_fails_on_blank_required_field(self, tmp_path):
        csv = tmp_path / "SUMMARY.csv"
        content = "\n".join(
            f"{f},{'' if f == 'customer_name' else 'val'}" for f in REQUIRED_SUMMARY_FIELDS
        )
        csv.write_text(content)
        with pytest.raises(ParseError, match="customer_name"):
            parse_summary(csv)

    def test_fails_on_unsupported_extension(self, tmp_path):
        f = tmp_path / "SUMMARY.json"
        f.write_text("{}")
        with pytest.raises(ParseError, match="Unsupported"):
            parse_summary(f)

    def test_parses_txt_with_colon_separator(self, tmp_path):
        txt = tmp_path / "SUMMARY.txt"
        fields = {f: "val" for f in REQUIRED_SUMMARY_FIELDS}
        content = "\n".join(f"{k}: {v}" for k, v in fields.items())
        txt.write_text(content)
        result = parse_summary(txt)
        assert result["customer_name"] == "val"

    def test_parses_txt_with_equals_separator(self, tmp_path):
        txt = tmp_path / "SUMMARY.txt"
        fields = {f: "val" for f in REQUIRED_SUMMARY_FIELDS}
        content = "\n".join(f"{k}={v}" for k, v in fields.items())
        txt.write_text(content)
        result = parse_summary(txt)
        assert result["conclusion"] == "val"

    def test_ignores_comment_lines(self, tmp_path):
        txt = tmp_path / "SUMMARY.txt"
        fields = {f: "val" for f in REQUIRED_SUMMARY_FIELDS}
        lines = ["# This is a comment"] + [f"{k}: {v}" for k, v in fields.items()]
        txt.write_text("\n".join(lines))
        result = parse_summary(txt)
        assert result["customer_name"] == "val"


class TestParseEquipment:
    def test_parses_fake_equipment_file(self):
        rows = parse_equipment(FAKE_EQUIPMENT)
        assert len(rows) >= 3
        assert rows[0]["equipment"] == "Shock Machine"
        assert rows[0]["make"] == "Lansmont"
        assert rows[0]["serial"] == "SN-DEMO-001"

    def test_all_equipment_columns_present(self):
        rows = parse_equipment(FAKE_EQUIPMENT)
        for row in rows:
            for col in EQUIPMENT_COLUMNS:
                assert col in row, f"Equipment row missing column: {col}"

    def test_fails_on_missing_column(self, tmp_path):
        csv = tmp_path / "EQUIPMENT.csv"
        csv.write_text("equipment,make,model\nShock Machine,Lansmont,XYZ\n")
        with pytest.raises(ParseError, match="missing required columns"):
            parse_equipment(csv)

    def test_fails_on_empty_file(self, tmp_path):
        csv = tmp_path / "EQUIPMENT.csv"
        csv.write_text(",".join(EQUIPMENT_COLUMNS) + "\n")
        with pytest.raises(ParseError, match="no data rows"):
            parse_equipment(csv)


class TestParseSequence:
    def test_parses_seq2_tip_over(self):
        data = parse_sequence(FAKE_SEQ2, "SEQ2_TIP_OVER")
        assert data["result"] == "Pass"
        assert data["test1_orientation"] == "Edge 2-3"
        assert data["test1_angle"] == "22"

    def test_fails_if_result_missing(self, tmp_path):
        csv = tmp_path / "SEQ2_TIP_OVER.csv"
        csv.write_text("height_inches,42\nweight_lbs,12.5\n")
        with pytest.raises(ParseError, match="result"):
            parse_sequence(csv, "SEQ2_TIP_OVER")

    def test_fails_if_result_blank(self, tmp_path):
        csv = tmp_path / "SEQ2_TIP_OVER.csv"
        csv.write_text("height_inches,42\nresult,\n")
        with pytest.raises(ParseError, match="blank"):
            parse_sequence(csv, "SEQ2_TIP_OVER")


class TestParseAll:
    def test_parse_all_returns_summary_fields(self):
        result = parse_all(LANSMONT_DIR, FAKE_SUMMARY)
        assert result["customer_name"] == "Customer_A"
        assert result["product_name"] == "Demo Product Sample"

    def test_parse_all_flattens_equipment(self):
        result = parse_all(LANSMONT_DIR, FAKE_SUMMARY)
        assert "equip_1_equipment" in result
        assert "equip_1_make" in result
        assert "equip_1_calibration_date" in result

    def test_parse_all_flattens_sequences(self):
        result = parse_all(LANSMONT_DIR, FAKE_SUMMARY)
        assert "seq2_result" in result
        assert result["seq2_result"] == "Pass"
        assert "seq5_vibration_intensity_grms" in result
        assert result["seq5_vibration_intensity_grms"] == "0.542"

    def test_parse_all_stores_private_equipment_rows(self):
        result = parse_all(LANSMONT_DIR, FAKE_SUMMARY)
        assert "_equipment_rows" in result
        assert isinstance(result["_equipment_rows"], list)

    def test_parse_all_stores_private_sequences(self):
        result = parse_all(LANSMONT_DIR, FAKE_SUMMARY)
        assert "_sequences" in result
        assert "SEQ2_TIP_OVER" in result["_sequences"]
