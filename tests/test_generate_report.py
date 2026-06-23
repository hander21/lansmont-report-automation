"""End-to-end test for ISTA 3B report generation using the fake sample folder."""

import json
import pytest
from pathlib import Path

from src.config import load_config
from src.validate_input import validate_input_folder, validate_parsed_fields
from src.parse_lansmont import parse_all
from src.organize_files import build_output_structure, copy_files
from src.generate_report import generate_draft_report, build_template_context
from src.audit_manifest import build_manifest, write_manifest

CONFIG_PATH = Path("config.yaml")
SAMPLE_FOLDER = Path("samples/fake_test_001")
LANSMONT_DIR = SAMPLE_FOLDER / "Lansmont"


@pytest.fixture
def cfg(tmp_path):
    base_cfg = load_config(CONFIG_PATH)
    base_cfg["output_folder"] = str(tmp_path / "output")
    return base_cfg


@pytest.fixture
def pipeline(cfg):
    """Run the full ISTA 3B pipeline and return all artifacts."""
    discovered = validate_input_folder(SAMPLE_FOLDER, cfg)
    parsed = parse_all(LANSMONT_DIR, discovered["summary_file"])
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
        name = pipeline["report_path"].name
        assert "Customer" in name or "CustomerA" in name

    def test_charts_copied_to_output(self, pipeline):
        assert len(pipeline["copied"]["charts"]) >= 5

    def test_pre_photos_copied(self, pipeline):
        assert len(pipeline["copied"]["pre_test_photos"]) >= 1

    def test_post_photos_copied(self, pipeline):
        assert len(pipeline["copied"]["post_test_photos"]) >= 1

    def test_accel_photos_copied(self, pipeline):
        assert len(pipeline["copied"]["accel_photos"]) >= 1

    def test_seq_photos_copied(self, pipeline):
        assert "Seq2_Tip_Over" in pipeline["copied"]["seq_photos"]
        assert len(pipeline["copied"]["seq_photos"]["Seq2_Tip_Over"]) >= 1

    def test_images_inserted_in_report(self, pipeline):
        assert len(pipeline["inserted_images"]) >= 1

    def test_parsed_has_sequence_fields(self, pipeline):
        parsed = pipeline["parsed"]
        assert "seq2_result" in parsed
        assert parsed["seq2_result"] == "Pass"
        assert "seq5_peak_g" in parsed

    def test_parsed_has_equipment_rows(self, pipeline):
        parsed = pipeline["parsed"]
        assert "_equipment_rows" in parsed
        assert len(parsed["_equipment_rows"]) >= 3


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
            assert "source_file" in entry, f"No source_file for: {field}"
            assert "source_field" in entry, f"No source_field for: {field}"
            assert "value" in entry, f"No value for: {field}"

    def test_manifest_records_correct_customer(self, pipeline):
        assert pipeline["manifest"]["customer_name"] == "Customer_A"

    def test_manifest_records_correct_project(self, pipeline):
        assert pipeline["manifest"]["project_number"] == "SAMPLE-123"

    def test_manifest_generated_at_is_present(self, pipeline):
        assert "generated_at" in pipeline["manifest"]
        assert pipeline["manifest"]["generated_at"]

    def test_manifest_includes_equipment_values(self, pipeline):
        values = pipeline["manifest"]["values_used"]
        assert any(k.startswith("equip_") for k in values)

    def test_manifest_includes_sequence_values(self, pipeline):
        values = pipeline["manifest"]["values_used"]
        assert any("seq2_tip_over" in k.lower() or "seq2" in k for k in values)


class TestTemplateContext:
    def test_draft_watermark_present(self):
        cfg = load_config(CONFIG_PATH)
        parsed = {
            "customer_name": "Customer_A",
            "project_number": "SAMPLE-123",
            "test_date": "2026-06-22",
            "test_type": "ISTA 3B",
            "product_name": "Demo",
            "packaging_description": "Demo box",
            "weight_lbs": "12.5",
            "dimensions_lwh": "18x12x10",
            "quantity": "6",
            "test_standard": "ISTA 3B-2021",
            "conclusion": "Pass",
        }
        ctx = build_template_context(parsed, cfg)
        assert "DRAFT" in ctx["draft_watermark"]

    def test_all_parsed_values_in_context(self):
        cfg = load_config(CONFIG_PATH)
        parsed = {
            "customer_name": "Customer_A",
            "project_number": "SAMPLE-123",
            "test_date": "2026-06-22",
            "test_type": "ISTA 3B",
            "product_name": "Demo Product",
            "packaging_description": "Demo box",
            "weight_lbs": "12.5",
            "dimensions_lwh": "18x12x10",
            "quantity": "6",
            "test_standard": "ISTA 3B-2021",
            "conclusion": "Pass",
            "seq2_result": "Pass",
        }
        ctx = build_template_context(parsed, cfg)
        assert ctx["customer_name"] == "Customer_A"
        assert ctx["weight_lbs"] == "12.5"
        assert ctx["seq2_result"] == "Pass"
