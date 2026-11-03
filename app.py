import contextlib
import io
import os
import shutil
import tempfile
from pathlib import Path

from flask import Flask, jsonify, render_template, request, send_file
from werkzeug.utils import secure_filename

from converter import convert

BASE_DIR = Path(__file__).resolve().parent
SAMPLE_PDF = BASE_DIR / "samples" / "sample_statement.pdf"
MAX_FILE_MB = int(os.environ.get("MAX_FILE_MB", "10"))

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = MAX_FILE_MB * 1024 * 1024


def wants_json() -> bool:
    return "application/json" in request.headers.get("Accept", "")


def error_response(message: str, status: int):
    if wants_json():
        return jsonify({"error": message}), status
    return render_template("index.html", error=message), status


@app.after_request
def add_security_headers(response):
    response.headers["X-Content-Type-Options"] = "nosniff"
    response.headers["Referrer-Policy"] = "same-origin"
    response.headers["X-Frame-Options"] = "DENY"
    response.headers["Content-Security-Policy"] = (
        "default-src 'self'; style-src 'self'; script-src 'self'; "
        "img-src 'self' data:; form-action 'self'; base-uri 'self'; "
        "frame-ancestors 'none'"
    )
    return response


@app.get("/")
def index():
    return render_template("index.html", error=None)


@app.get("/health")
def health():
    return jsonify({"status": "ok", "service": "SimpoSheet"})


@app.get("/sample-statement")
def sample_statement():
    return send_file(
        SAMPLE_PDF,
        mimetype="application/pdf",
        as_attachment=True,
        download_name="simposheet-sample-statement.pdf",
    )


@app.post("/convert")
def convert_upload():
    upload = request.files.get("document")
    if upload is None or not upload.filename:
        return error_response("Please choose a PDF file to convert.", 400)

    original_name = secure_filename(upload.filename)
    if not original_name.lower().endswith(".pdf"):
        return error_response("SimpoSheet currently accepts PDF files only.", 400)

    header = upload.stream.read(5)
    upload.stream.seek(0)
    if header != b"%PDF-":
        return error_response("That file does not look like a valid PDF.", 400)

    workdir = Path(tempfile.mkdtemp(prefix="simposheet-"))
    input_path = workdir / "input.pdf"
    output_name = f"{Path(original_name).stem}-SimpoSheet.xlsx"
    output_path = workdir / "output.xlsx"

    try:
        upload.save(input_path)

        # The prototype converter prints a result summary. Suppress it so
        # document details never enter the web server's request logs.
        with contextlib.redirect_stdout(io.StringIO()):
            convert(str(input_path), str(output_path))

        if not output_path.exists() or output_path.stat().st_size == 0:
            shutil.rmtree(workdir, ignore_errors=True)
            return error_response(
                "We could not extract a usable workbook from that PDF yet. "
                "Text-based bank statements, receipts, and supported bills work best.",
                422,
            )

        response = send_file(
            output_path,
            mimetype=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            as_attachment=True,
            download_name=output_name,
        )
        response.call_on_close(lambda: shutil.rmtree(workdir, ignore_errors=True))
        response.headers["X-SimpoSheet-Privacy"] = "temporary-file-deleted-after-response"
        return response
    except Exception:
        shutil.rmtree(workdir, ignore_errors=True)
        app.logger.exception("Conversion failed")
        return error_response(
            "Something went wrong while converting that file. Please try a "
            "text-based PDF or a smaller file.",
            500,
        )


@app.errorhandler(413)
def too_large(_error):
    return error_response(
        f"That file is too large. The current limit is {MAX_FILE_MB} MB.",
        413,
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "5000"))
    app.run(host="0.0.0.0", port=port)
