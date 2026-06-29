"""
Lansmont Report Automation — CLI entry point.

Usage:
    python -m src.main [--input INPUT_FOLDER] [--config CONFIG_PATH]

The tool reads a completed test folder, validates all required files,
organizes them into the output structure, generates a DRAFT Word report,
and writes a manifest.json audit file.

No real customer data, product photos, or Lansmont exports should be
processed during development. See CLAUDE.md for confidentiality rules.
"""

import argparse
import logging
import sys
from pathlib import Path

from src.config import load_config
from src.utils import setup_logging
from src.validate_input import validate_input_folder, validate_template, validate_parsed_fields, ValidationError
from src.parse_lansmont import parse_summary, parse_all, ParseError
from src.organize_files import build_output_structure, copy_files
from src.generate_report import generate_draft_report
from src.audit_manifest import build_manifest, write_manifest

logger = logging.getLogger(__name__)


def parse_args(argv=None):
    parser = argparse.ArgumentParser(
        description="Generate a DRAFT packaging test report from a Lansmont test folder."
    )
    parser.add_argument(
        "--input",
        type=Path,
        help="Path to the test input folder (overrides config.yaml input_folder).",
    )
    parser.add_argument(
        "--config",
        type=Path,
        default=Path("config.yaml"),
        help="Path to config.yaml (default: config.yaml in current directory).",
    )
    parser.add_argument(
        "--log-level",
        default="INFO",
        choices=["DEBUG", "INFO", "WARNING", "ERROR"],
        help="Logging verbosity.",
    )
    return parser.parse_args(argv)


def run(argv=None) -> int:
    args = parse_args(argv)
    setup_logging(args.log_level)

    logger.info("=" * 60)
    logger.info("Lansmont Report Automation — DRAFT mode")
    logger.info("=" * 60)

    # Load config
    try:
        cfg = load_config(args.config)
    except (FileNotFoundError, ValueError) as e:
        logger.error("CONFIG ERROR: %s", e)
        return 1

    input_folder = args.input or Path(cfg["input_folder"])

    logger.info("Input folder : %s", input_folder.resolve())
    logger.info("Output base  : %s", cfg['output_folder'])

    # Step 1: Validate input folder (flexible scan — any folder structure accepted)
    try:
        discovered = validate_input_folder(input_folder, cfg)
    except ValidationError as e:
        logger.error("INPUT VALIDATION FAILED:\n%s", e)
        return 1

    # Step 2: Parse all input data (summary + equipment + sequences)
    if not discovered.get("summary_file"):
        logger.error("No summary file was discovered. Cannot continue.")
        return 1

    try:
        # Use the folder that contains the summary file as the data root
        data_folder = discovered["summary_file"].parent
        parsed = parse_all(data_folder, discovered["summary_file"])
    except ParseError as e:
        logger.error("PARSE ERROR: %s", e)
        return 1

    # Step 4: Validate that all required fields are present in parsed data
    try:
        validate_parsed_fields(parsed)
    except ValidationError as e:
        logger.error("FIELD VALIDATION FAILED: %s", e)
        return 1

    logger.info(
        "Test: %s | Customer: %s | Project: %s | Date: %s",
        parsed["test_type"],
        parsed["customer_name"],
        parsed["project_number"],
        parsed["test_date"],
    )

    # Step 4b: Select template based on test standard
    std = (parsed.get("test_standard", "") + " " + parsed.get("test_type", "")).upper()
    if "2A" in std:
        template_path = Path(cfg.get("ista2a_template_path", cfg["template_path"]))
        logger.info("Test standard detected as ISTA 2A — using 2A template")
    elif "3A" in std:
        template_path = Path(cfg.get("ista3a_template_path", cfg["template_path"]))
        logger.info("Test standard detected as ISTA 3A — using 3A template")
    elif "SIMPLIFIED" in std or "CUSTOM" in std:
        template_path = Path(cfg.get("ista_simplified_template_path", cfg["template_path"]))
        logger.info("Test standard detected as Simplified/Custom — using Simplified template")
    else:
        template_path = Path(cfg.get("ista3b_template_path", cfg["template_path"]))
        logger.info("Test standard detected as ISTA 3B — using 3B template")

    cfg["template_path"] = str(template_path)

    try:
        validate_template(template_path)
    except ValidationError as e:
        logger.error("TEMPLATE ERROR: %s", e)
        return 1

    logger.info("Template     : %s", template_path.resolve())

    # Step 5: Organize files into output structure
    output_paths = build_output_structure(parsed, cfg)
    copied_files = copy_files(discovered, output_paths)

    # Step 6: Generate draft report
    try:
        generated_report, inserted_images = generate_draft_report(
            parsed, discovered, output_paths, cfg
        )
    except Exception as e:
        logger.error("REPORT GENERATION FAILED: %s", e)
        return 1

    # Step 7: Write manifest
    warnings = discovered.get("warnings", [])
    manifest = build_manifest(
        parsed=parsed,
        discovered=discovered,
        output_paths=output_paths,
        copied_files=copied_files,
        inserted_images=inserted_images,
        generated_report=generated_report,
        input_folder=input_folder,
        warnings=warnings,
    )
    manifest_path = write_manifest(manifest, output_paths)

    # Done
    logger.info("=" * 60)
    logger.info("DRAFT REPORT GENERATED SUCCESSFULLY")
    logger.info("  Report  : %s", generated_report)
    logger.info("  Manifest: %s", manifest_path)
    logger.info("")
    logger.info("ACTION REQUIRED: A qualified test engineer must review")
    logger.info("  the DRAFT report before it is sent to any customer.")
    logger.info("=" * 60)

    return 0


if __name__ == "__main__":
    sys.exit(run())
