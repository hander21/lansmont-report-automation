"""End-to-end test for report generation using the fake sample folder."""

import json
import pytest
from pathlib import Path

from src.config import load_config
from src.validate_input import validate_input_folder, validate_parsed_fields
from src.parse_lansmont import parse_summary
from src.organize_files import build_output_structure, copy_files
from src.generate_report import generate_draft_report, build_template_context
from src.audit_manifest import build_manifest, write_manifest

CONFIG_PATH = Path("config.yaml")
SAMPLE_FOLDER = Path("samples/fake_test_001")


@pytest.fixture
def cfg(tmp_path):
    base_cfg = load_config(CONFIG_PATH)
    base_cfg["output_folder"] = str(tmp_path / "output")
    return base_cfg


@pytest.fixture
def pipeline(cfg):
    """Run the full pipeline and return all artifacts."""
    discovered = validate_input_folder(SAMPLE_FOLDER, cfg)
    parsed = parse_summary(discovered["summary_file"])
    validate_parsed_fields(parsed)
    output_paths = build_output_structure(parsed, cfg)
    copied = copy_files(discovered, output_paths)
    report_path, inserted_images = generate_draft_report(parsed, discovered, output_paths, cfg)
    manifest = build_manifest(
        parsed=parsed,
        discovered=discovered,
        output_paths=output_paths,
        copied_files=copied,
        inserted_images=inserted_images,
        generated_report=report_path,
        input_folder=SAMPLE_FOLDER,
        warnings=discovered.get("warnings", []),
    )
    manifest_path = write_manifest(manifest, output_paths)
    return {
        "parsed": parsed,
        "output_paths": output_paths,
        "copied": copied,
        "report_path": report_path,
        "inserted_images": inserted_images,
        "manifest": manifest,
        "manifest_path": manifest_path,
    }


class TestReportGeneration:
    def test_report_file_created(self, pipeline):
        assert pipeline["report_path"].exists()

    def test_report_filename_starts_with_draft(self, pipeline):
        assert pipeline["report_path"].name.startswith("DRAFT_")

    def test_report_is_docx(self, pipeline):
        assert pipeline["report_path"].suffix == ".docx"

    def test_report_filename_contains_project_number(self, pipeline):
        assert "SAMPLE-123" in pipeline["report_path"].name

    def test_report_filename_contains_customer(self, pipeline):
        assert "Customer" in pipeline["report_path"].name or "CustomerA" in pipeline["report_path"].name

    def test_charts_copied_to_output(self, pipeline):
        assert len(pipeline["copied"]["charts"]) >= 2

    def test_pre_photos_copied(self, pipeline):
        assert len(pipeline["copied"]["pre_test_photos"]) >= 1

    def test_post_photos_copied(self, pipeline):
        assert len(pipeline["copied"]["post_test_photos"]) >= 1


class TestManifest:
    def test_manifest_file_exists(self, pipeline):
        assert pipeline["manifest_path"].exists()

    def test_manifest_is_valid_json(self, pipeline):
        with open(pipeline["manifest_path"]) as f:
            data = json.load(f)
        assert isinstance(data, dict)

    def test_manifest_has_required_keys(self, pipeline):
        m = pipeline["manifest"]
        for key in [
            "customer_name", "project_number", "test_date", "test_type",
            "input_folder", "output_folder", "source_files", "values_used",
            "generated_report", "report_status", "human_review_required",
        ]:
            assert key in m, f"Manifest missing key: {key}"

    def test_manifest_marks_report_as_draft(self, pipeline):
        assert pipeline["manifest"]["report_status"] == "DRAFT"

    def test_manifest_requires_human_review(self, pipeline):
        assert pipeline["manifest"]["human_review_required"] is True

    def test_manifest_values_have_source_traceability(self, pipeline):
        for field, entry in pipeline["manifest"]["values_used"].items():
            assert "source_file" in entry, f"No source_file for field: {field}"
            assert "source_field" in entry, f"No source_field for field: {field}"
            assert "value" in entry, f"No value for field: {field}"

    def test_manifest_records_correct_customer(self, pipeline):
        assert pipeline["manifest"]["customer_name"] == "Customer_A"

    def test_manifest_records_correct_project(self, pipeline):
        assert pipeline["manifest"]["project_number"] == "SAMPLE-123"

    def test_manifest_generated_at_is_present(self, pipeline):
        assert "generated_at" in pipeline["manifest"]
        assert pipeline["manifest"]["generated_at"]  # not empty


class TestTemplateContext:
    def test_draft_watermark_present(self):
        cfg = load_config(CONFIG_PATH)
        parsed = {"customer_name": "Customer_A", "project_number": "SAMPLE-123",
                  "test_date": "2026-06-22", "test_type": "Vibration Table",
                  "package_description": "Demo", "test_standard": "Std-1",
                  "peak_g": "38.4", "duration": "120 seconds",
                  "axis": "Z", "result": "Pass"}
        ctx = build_template_context(parsed, cfg)
        assert "DRAFT" in ctx["draft_watermark"]

    def test_all_parsed_values_in_context(self):
        cfg = load_config(CONFIG_PATH)
        parsed = {"customer_name": "Customer_A", "project_number": "SAMPLE-123",
                  "test_date": "2026-06-22", "test_type": "Vibration Table",
                  "package_description": "Demo", "test_standard": "Std-1",
                  "peak_g": "38.4", "duration": "120 seconds",
                  "axis": "Z", "result": "Pass"}
        ctx = build_template_context(parsed, cfg)
        assert ctx["customer_name"] == "Customer_A"
        assert ctx["peak_g"] == "38.4"
