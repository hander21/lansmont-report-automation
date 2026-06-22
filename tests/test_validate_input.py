"""Tests for src/validate_input.py using the fake sample folder."""

import pytest
from pathlib import Path

from src.config import load_config
from src.validate_input import (
    validate_input_folder,
    validate_template,
    validate_parsed_fields,
    ValidationError,
    REQUIRED_SUMMARY_FIELDS,
)

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
        assert len(discovered["charts"]) >= 2
        assert len(discovered["pre_test_photos"]) >= 1
        assert len(discovered["post_test_photos"]) >= 1

    def test_fails_if_folder_missing(self, cfg):
        with pytest.raises(ValidationError, match="does not exist"):
            validate_input_folder(Path("nonexistent/folder"), cfg)

    def test_fails_if_path_is_file(self, cfg, tmp_path):
        f = tmp_path / "not_a_dir.txt"
        f.write_text("hello")
        with pytest.raises(ValidationError, match="not a directory"):
            validate_input_folder(f, cfg)

    def test_fails_if_no_summary_file(self, cfg, tmp_path):
        (tmp_path / "Lansmont").mkdir()
        (tmp_path / "Photos" / "Pre-Test").mkdir(parents=True)
        (tmp_path / "Photos" / "Post-Test").mkdir(parents=True)
        with pytest.raises(ValidationError, match="summary file"):
            validate_input_folder(tmp_path, cfg)

    def test_fails_if_no_charts(self, cfg, tmp_path):
        lansmont = tmp_path / "Lansmont"
        lansmont.mkdir()
        (lansmont / "SUMMARY.csv").write_text("customer_name,Test\n")
        (tmp_path / "Photos" / "Pre-Test").mkdir(parents=True)
        (tmp_path / "Photos" / "Post-Test").mkdir(parents=True)
        with pytest.raises(ValidationError, match="[Cc]hart"):
            validate_input_folder(tmp_path, cfg)

    def test_fails_if_no_pre_photos(self, cfg, tmp_path):
        lansmont = tmp_path / "Lansmont"
        lansmont.mkdir()
        (lansmont / "SUMMARY.csv").write_text("customer_name,Test\n")
        (lansmont / "CHART_VIB.png").write_bytes(b"\x89PNG\r\n")
        (tmp_path / "Photos" / "Pre-Test").mkdir(parents=True)
        (tmp_path / "Photos" / "Post-Test").mkdir(parents=True)
        with pytest.raises(ValidationError, match="[Pp]re.test"):
            validate_input_folder(tmp_path, cfg)

    def test_summary_file_is_pathlib_path(self, cfg):
        discovered = validate_input_folder(SAMPLE_FOLDER, cfg)
        assert isinstance(discovered["summary_file"], Path)

    def test_charts_are_pathlib_paths(self, cfg):
        discovered = validate_input_folder(SAMPLE_FOLDER, cfg)
        for c in discovered["charts"]:
            assert isinstance(c, Path)


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

    def test_fails_if_field_missing(self):
        incomplete = {f: "val" for f in REQUIRED_SUMMARY_FIELDS if f != "peak_g"}
        with pytest.raises(ValidationError, match="peak_g"):
            validate_parsed_fields(incomplete)

    def test_fails_if_field_empty_string(self):
        fields = {f: "val" for f in REQUIRED_SUMMARY_FIELDS}
        fields["result"] = ""
        with pytest.raises(ValidationError, match="result"):
            validate_parsed_fields(fields)

    def test_fails_if_field_whitespace_only(self):
        fields = {f: "val" for f in REQUIRED_SUMMARY_FIELDS}
        fields["customer_name"] = "   "
        with pytest.raises(ValidationError, match="customer_name"):
            validate_parsed_fields(fields)

    def test_all_required_fields_listed(self):
        assert "peak_g" in REQUIRED_SUMMARY_FIELDS
        assert "customer_name" in REQUIRED_SUMMARY_FIELDS
        assert "result" in REQUIRED_SUMMARY_FIELDS
