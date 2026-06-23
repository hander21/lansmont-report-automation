"""
Local web UI for Lansmont Report Automation.
Run with: py -m ui.app
Opens at http://localhost:5000

Real data never leaves this machine. No network calls are made.
"""

import io
import logging
import sys
import tempfile
import threading
import zipfile
from pathlib import Path

from flask import Flask, render_template, request, jsonify, send_file

ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(ROOT))

from src.config import load_config
from src.validate_input import validate_input_folder, validate_template, validate_parsed_fields, ValidationError
from src.parse_lansmont import parse_summary, ParseError
from src.organize_files import build_output_structure, copy_files
from src.generate_report import generate_draft_report
from src.audit_manifest import build_manifest, write_manifest

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 500 * 1024 * 1024  # 500 MB max upload
log = logging.getLogger(__name__)

CONFIG_PATH = ROOT / "config.yaml"
TEMPLATE_PATH = ROOT / "templates" / "fake_report_template.docx"

# Holds last-generated output so download endpoints can serve it
_last_output: dict = {}


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/upload", methods=["POST"])
def upload():
    """
    Accept a ZIP file, extract it, validate, run the full pipeline,
    and return metadata. Downloads are served via /api/download/*.
    """
    global _last_output

    if "file" not in request.files:
        return jsonify({"ok": False, "error": "No file uploaded."}), 400

    f = request.files["file"]
    if not f.filename.lower().endswith(".zip"):
        return jsonify({"ok": False, "error": "Please upload a .zip file."}), 400

    # Extract ZIP into a temp folder
    tmp_in = tempfile.mkdtemp(prefix="lansmont_in_")
    tmp_out = tempfile.mkdtemp(prefix="lansmont_out_")

    try:
        with zipfile.ZipFile(f.stream) as zf:
            zf.extractall(tmp_in)
    except zipfile.BadZipFile:
        return jsonify({"ok": False, "error": "The uploaded file is not a valid ZIP."}), 400

    # The ZIP may contain a top-level folder or drop files directly — find the root
    input_folder = _find_input_root(Path(tmp_in))

    cfg = load_config(CONFIG_PATH)
    cfg["output_folder"] = tmp_out
    cfg["template_path"] = str(TEMPLATE_PATH)

    try:
        validate_template(TEMPLATE_PATH)
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

        _last_output = {
            "report_path": report_path,
            "manifest_path": manifest_path,
            "output_root": output_paths["report_root"],
            "parsed": parsed,
            "warnings": discovered.get("warnings", []),
        }

        return jsonify({
            "ok": True,
            "customer_name": parsed.get("customer_name"),
            "project_number": parsed.get("project_number"),
            "test_date": parsed.get("test_date"),
            "test_type": parsed.get("test_type"),
            "report_name": report_path.name,
            "warnings": discovered.get("warnings", []),
        })

    except (ValidationError, ParseError) as e:
        return jsonify({"ok": False, "error": str(e)}), 400
    except Exception as e:
        log.exception("Unexpected error")
        return jsonify({"ok": False, "error": f"Unexpected error: {e}"}), 500


@app.route("/api/download/docx")
def download_docx():
    if not _last_output or not _last_output.get("report_path"):
        return "No report generated yet.", 404
    path = _last_output["report_path"]
    return send_file(path, as_attachment=True, download_name=path.name)


@app.route("/api/download/manifest")
def download_manifest():
    if not _last_output or not _last_output.get("manifest_path"):
        return "No manifest generated yet.", 404
    path = _last_output["manifest_path"]
    return send_file(path, as_attachment=True, download_name="manifest.json")


@app.route("/api/download/zip")
def download_zip():
    """Bundle the entire organized output folder into a ZIP for download."""
    if not _last_output or not _last_output.get("output_root"):
        return "No output generated yet.", 404

    output_root = Path(_last_output["output_root"])
    parsed = _last_output.get("parsed", {})
    zip_name = (
        f"DRAFT_{parsed.get('customer_name','report')}_"
        f"{parsed.get('project_number','')}_"
        f"{parsed.get('test_date','')}.zip"
    ).replace(" ", "")

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for file in output_root.rglob("*"):
            if file.is_file():
                zf.write(file, file.relative_to(output_root.parent))
    buf.seek(0)

    return send_file(
        buf,
        as_attachment=True,
        download_name=zip_name,
        mimetype="application/zip",
    )


def _find_input_root(extracted: Path) -> Path:
    """
    If the ZIP had a single top-level folder, return that folder.
    Otherwise return the extraction root itself.
    """
    children = [c for c in extracted.iterdir()]
    if len(children) == 1 and children[0].is_dir():
        return children[0]
    return extracted


def main():
    logging.basicConfig(level=logging.INFO,
                        format="%(asctime)s [%(levelname)s] %(message)s")
    import webbrowser
    threading.Timer(1.2, lambda: webbrowser.open("http://localhost:5000")).start()
    print("\n  Lansmont Report Automation — Local UI")
    print("  http://localhost:5000")
    print("  Press Ctrl+C to stop.\n")
    app.run(host="127.0.0.1", port=5000, debug=False)


if __name__ == "__main__":
    main()
