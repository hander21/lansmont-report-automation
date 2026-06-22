"""
Local web UI for Lansmont Report Automation.
Run with: python -m ui.app
Opens at http://localhost:5000

Real data never leaves this machine. No network calls are made.
"""

import json
import logging
import sys
import threading
from pathlib import Path

from flask import Flask, render_template, request, jsonify, send_file

# Add project root to path so src.* imports work
ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(ROOT))

from src.config import load_config
from src.validate_input import validate_input_folder, validate_template, validate_parsed_fields, ValidationError
from src.parse_lansmont import parse_summary, ParseError
from src.organize_files import build_output_structure, copy_files
from src.generate_report import generate_draft_report
from src.audit_manifest import build_manifest, write_manifest

app = Flask(__name__)
log = logging.getLogger(__name__)

CONFIG_PATH = ROOT / "config.yaml"


@app.route("/")
def index():
    cfg = load_config(CONFIG_PATH)
    return render_template(
        "index.html",
        default_input=cfg.get("input_folder", ""),
        default_output=cfg.get("output_folder", ""),
        default_template=cfg.get("template_path", ""),
    )


@app.route("/api/validate", methods=["POST"])
def validate():
    """Validate the input folder without generating a report."""
    data = request.get_json()
    input_folder = Path(data.get("input_folder", ""))
    template_path = Path(data.get("template_path", ""))

    cfg = _build_runtime_cfg(data)
    issues = []
    warnings = []
    fields = {}

    try:
        validate_template(template_path)
    except ValidationError as e:
        issues.append(str(e))

    try:
        discovered = validate_input_folder(input_folder, cfg)
        warnings.extend(discovered.get("warnings", []))

        if discovered.get("summary_file"):
            try:
                parsed = parse_summary(discovered["summary_file"])
                validate_parsed_fields(parsed)
                fields = parsed
            except (ParseError, ValidationError) as e:
                issues.append(str(e))
    except ValidationError as e:
        issues.append(str(e))

    return jsonify({
        "ok": len(issues) == 0,
        "errors": issues,
        "warnings": warnings,
        "fields": fields,
    })


@app.route("/api/generate", methods=["POST"])
def generate():
    """Run the full pipeline and return the output path."""
    data = request.get_json()
    input_folder = Path(data.get("input_folder", ""))
    template_path = Path(data.get("template_path", ""))
    cfg = _build_runtime_cfg(data)

    try:
        validate_template(template_path)
        discovered = validate_input_folder(input_folder, cfg)
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
            input_folder=input_folder,
            warnings=discovered.get("warnings", []),
        )
        manifest_path = write_manifest(manifest, output_paths)

        return jsonify({
            "ok": True,
            "report": str(report_path),
            "manifest": str(manifest_path),
            "output_folder": str(output_paths["report_root"]),
            "customer_name": parsed.get("customer_name"),
            "project_number": parsed.get("project_number"),
            "test_date": parsed.get("test_date"),
            "test_type": parsed.get("test_type"),
            "warnings": discovered.get("warnings", []),
        })

    except (ValidationError, ParseError) as e:
        return jsonify({"ok": False, "error": str(e)}), 400
    except Exception as e:
        log.exception("Unexpected error during report generation")
        return jsonify({"ok": False, "error": f"Unexpected error: {e}"}), 500


@app.route("/api/download")
def download():
    """Download the generated report file."""
    path = Path(request.args.get("path", ""))
    if not path.exists() or not path.is_file():
        return jsonify({"error": "File not found"}), 404
    if path.suffix.lower() not in (".docx", ".json"):
        return jsonify({"error": "Invalid file type"}), 400
    return send_file(path, as_attachment=True)


def _build_runtime_cfg(form_data: dict) -> dict:
    """Merge form inputs over the base config."""
    cfg = load_config(CONFIG_PATH)
    if form_data.get("input_folder"):
        cfg["input_folder"] = form_data["input_folder"]
    if form_data.get("output_folder"):
        cfg["output_folder"] = form_data["output_folder"]
    if form_data.get("template_path"):
        cfg["template_path"] = form_data["template_path"]
    return cfg


def main():
    logging.basicConfig(level=logging.INFO,
                        format="%(asctime)s [%(levelname)s] %(message)s")
    import webbrowser
    # Open browser after a short delay so Flask has time to start
    threading.Timer(1.2, lambda: webbrowser.open("http://localhost:5000")).start()
    print("\n  Lansmont Report Automation — Local UI")
    print("  http://localhost:5000")
    print("  Press Ctrl+C to stop.\n")
    app.run(host="127.0.0.1", port=5000, debug=False)


if __name__ == "__main__":
    main()
