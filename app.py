"""
Research Paper Metadata Dashboard
=================================
Web UI on top of extract_metadata.py:
  - drag & drop PDFs -> extract title / authors / emails
  - results table, one row per author email
  - download everything as a formatted Excel (.xlsx)
  - upload an existing metadata Excel to validate it

Run:  python app.py   ->  http://127.0.0.1:5075
"""

import io
import os
import tempfile
import uuid

from flask import Flask, jsonify, render_template, request, send_file
from werkzeug.utils import secure_filename

from extract_metadata import HEADERS, check_excel, process_pdf, write_excel

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024  # 200 MB per request

# extracted records for the current session, keyed by filename (insertion order kept)
RECORDS = {}


@app.get("/")
def index():
    return render_template("index.html")


@app.post("/api/extract")
def api_extract():
    files = request.files.getlist("pdfs")
    if not files:
        return jsonify({"error": "no files uploaded"}), 400

    results = []
    with tempfile.TemporaryDirectory() as tmp:
        for f in files:
            name = secure_filename(os.path.basename(f.filename or ""))
            if not name.lower().endswith(".pdf"):
                results.append({"file": f.filename, "title": "", "authors": [],
                                "emails": [], "error": "not a PDF"})
                continue
            path = os.path.join(tmp, f"{uuid.uuid4().hex}_{name}")
            f.save(path)
            rec = process_pdf(path)
            rec["file"] = name  # keep the original filename, not the temp one
            RECORDS[name] = rec
            results.append(rec)
    return jsonify({"results": results, "total_papers": len(RECORDS)})


@app.get("/api/records")
def api_records():
    return jsonify({"results": list(RECORDS.values()), "total_papers": len(RECORDS)})


@app.post("/api/clear")
def api_clear():
    RECORDS.clear()
    return jsonify({"ok": True})


@app.delete("/api/records/<name>")
def api_delete(name):
    RECORDS.pop(name, None)
    return jsonify({"ok": True, "total_papers": len(RECORDS)})


@app.get("/api/download")
def api_download():
    if not RECORDS:
        return jsonify({"error": "nothing extracted yet"}), 400
    tmp = os.path.join(tempfile.gettempdir(), f"metadata_{uuid.uuid4().hex}.xlsx")
    write_excel(list(RECORDS.values()), tmp)
    with open(tmp, "rb") as fh:
        data = fh.read()
    os.remove(tmp)
    return send_file(io.BytesIO(data), as_attachment=True,
                     download_name="Research_Paper_Metadata.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.post("/api/check")
def api_check():
    f = request.files.get("xlsx")
    if not f:
        return jsonify({"error": "no file uploaded"}), 400
    name = secure_filename(os.path.basename(f.filename or "sheet.xlsx"))
    path = os.path.join(tempfile.gettempdir(), f"{uuid.uuid4().hex}_{name}")
    f.save(path)
    try:
        n_rows, n_papers, problems = check_excel(path)
    except Exception as exc:
        return jsonify({"error": f"could not read Excel: {exc}"}), 400
    finally:
        os.remove(path)
    return jsonify({"file": f.filename, "rows": n_rows, "papers": n_papers,
                    "problems": problems})


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5075, debug=False)
