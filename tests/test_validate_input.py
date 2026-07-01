"""Tests for src/validate_input.py using the fake ISTA 3B sample folder."""

import pytest
from pathlib import Path

from src.config import load_config
from src.validate_input import (
    validate_input_folder,
    validate_template,
    validate_parsed_fields,
    ValidationError,
)
from src.parse_lansmont import REQUIRED_SUMMARY_FIELDS

CONFIG_PATH = Path("config.yaml")
SAMPLE_FOLDER = Path("samples/fake_test_001")
TEMPLATE_PATH = Path("templates/fake_report_template.docx")


@pytest.fixture
def cfg():
    return load_config(CONFIG_PATH)


class TestValidateInputFolder:
    def test_passes_with_valid_sample(self, cfg):
        discovered = validate_input_folder(SAMPLE_FOLDER, cfg)
        assert discovered["summary_file"] is not None
        assert len(discovered["charts"]) >= 5
        assert len(discovered["pre_test_photos"]) >= 1
        assert len(discovered["post_test_photos"]) >= 1

    def test_discovers_accelerometer_photos(self, cfg):
        discovered = validate_input_folder(SAMPLE_FOLDER, cfg)
        assert len(discovered["accel_photos"]) >= 1

    def test_discovers_seq_photos(self, cfg):
        discovered = validate_input_folder(SAMPLE_FOLDER, cfg)
        assert "Seq2_Tip_Over" in discovered["seq_photos"]
        assert len(discovered["seq_photos"]["Seq2_Tip_Over"]) >= 1

    def test_fails_if_folder_missing(self, cfg):
        with pytest.raises(ValidationError, match="does not exist"):
            validate_input_folder(Path("nonexistent/folder"), cfg)

    def test_fails_if_path_is_file(self, cfg, tmp_path):
        f = tmp_path / "not_a_dir.txt"
        f.write_text("hello")
        with pytest.raises(ValidationError, match="not a directory"):
            validate_input_folder(f, cfg)

    def test_no_summary_file_produces_warning_not_error(self, cfg, tmp_path):
        """Missing summary is now a warning — report fields will be blank."""
        (tmp_path / "Lansmont").mkdir()
        discovered = validate_input_folder(tmp_path, cfg)
        assert discovered["summary_file"] is None
        assert any("summary" in w.lower() for w in discovered["warnings"])

    def test_no_charts_produces_warning_not_error(self, cfg, tmp_path):
        """Missing charts are now a warning — slots remain as placeholders."""
        lansmont = tmp_path / "Lansmont"
        lansmont.mkdir()
        (lansmont / "SUMMARY.csv").write_text("customer_name,Test\n")
        discovered = validate_input_folder(tmp_path, cfg)
        assert discovered["charts"] == []
        assert any("chart" in w.lower() for w in discovered["warnings"])

    def test_no_pre_photos_produces_warning_not_error(self, cfg, tmp_path):
        """Photos are optional — no photos should warn, not hard-fail."""
        lansmont = tmp_path / "Lansmont"
        lansmont.mkdir()
        (lansmont / "SUMMARY.csv").write_text("customer_name,Test\n")
        (lansmont / "CHART_VIB.png").write_bytes(b"\x89PNG\r\n")
        # No Photos folder at all — should succeed with warnings
        discovered = validate_input_folder(tmp_path, cfg)
        assert len(discovered["pre_test_photos"]) == 0
        assert any("pre" in w for w in discovered["warnings"])

    def test_summary_file_is_pathlib_path(self, cfg):
        discovered = validate_input_folder(SAMPLE_FOLDER, cfg)
        assert isinstance(discovered["summary_file"], Path)

    def test_charts_are_pathlib_paths(self, cfg):
        discovered = validate_input_folder(SAMPLE_FOLDER, cfg)
        for c in discovered["charts"]:
            assert isinstance(c, Path)

    def test_missing_photos_produce_warnings_not_errors(self, cfg, tmp_path):
        """All photos are optional — missing photos warn, never hard-fail."""
        lansmont = tmp_path / "Lansmont"
        lansmont.mkdir()
        (lansmont / "SUMMARY.csv").write_text("customer_name,Test\n")
        (lansmont / "CHART_VIB.png").write_bytes(b"\x89PNG\r\n")
        # No photos at all — should succeed with warnings for every bucket
        discovered = validate_input_folder(tmp_path, cfg)
        assert discovered["pre_test_photos"] == []
        assert discovered["post_test_photos"] == []
        assert len(discovered["warnings"]) >= 1


class TestValidateTemplate:
    def test_passes_with_real_template(self):
        validate_template(TEMPLATE_PATH)

    def test_fails_if_missing(self):
        with pytest.raises(ValidationError, match="not found"):
            validate_template(Path("nonexistent.docx"))

    def test_fails_if_not_docx(self, tmp_path):
        f = tmp_path / "report.pdf"
        f.write_bytes(b"%PDF")
        with pytest.raises(ValidationError, match=".docx"):
            validate_template(f)


class TestValidateParsedFields:
    def test_passes_with_complete_fields(self):
        complete = {f: "some_value" for f in REQUIRED_SUMMARY_FIELDS}
        validate_parsed_fields(complete)  # should not raise

    def test_incomplete_fields_do_not_raise(self):
        """Missing fields warn, not error — report will have blank placeholders."""
        incomplete = {f: "val" for f in REQUIRED_SUMMARY_FIELDS if f != "conclusion"}
        validate_parsed_fields(incomplete)  # should not raise

    def test_blank_fields_do_not_raise(self):
        """Blank fields warn, not error."""
        fields = {f: "val" for f in REQUIRED_SUMMARY_FIELDS}
        fields["customer_name"] = ""
        validate_parsed_fields(fields)  # should not raise

    def test_required_fields_include_ista3b_fields(self):
        assert "product_name" in REQUIRED_SUMMARY_FIELDS
        assert "weight_lbs" in REQUIRED_SUMMARY_FIELDS
        assert "conclusion" in REQUIRED_SUMMARY_FIELDS
        assert "customer_name" in REQUIRED_SUMMARY_FIELDS
