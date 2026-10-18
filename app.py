"""Local browser app for the PDF / P&ID tag extractor.

Run:  python app.py
The command-line tool is unchanged:  python pdf_tag_extractor_multi.py
"""

import json
import os
import shutil
import socket
import threading
import traceback
import uuid
import webbrowser
from datetime import datetime

from flask import Flask, jsonify, render_template, request, send_file

import pdf_tag_extractor_multi as eng

BASE = os.path.dirname(os.path.abspath(__file__))
WORK = os.path.join(BASE, "work")
PREVIEW_LIMIT = 200
MAX_JOBS = 4

app = Flask(__name__, template_folder=os.path.join(BASE, "templates"), static_folder=os.path.join(BASE, "static"))
app.config["MAX_CONTENT_LENGTH"] = 2 * 1024 * 1024 * 1024
app.json.sort_keys = False

JOBS = {}
LOCK = threading.Lock()

EXAMPLES = [
    ("tag no", "Every tag in the lists, each with its own serial no., model, part number, and reference file."),
    ("PIT-01043", "Exact tag. Also finds PIT-01043A, PIT-01043B, and similar suffixes."),
    ("PIT", "Every instrument whose function is PIT."),
    ("01043", "Every item containing 01043, such as PIT-01043 or TIT-01043."),
    ("PIT-010*", "Wildcard. * means anything, ? means one character."),
    ("866-E-3303B", "Equipment tag."),
    ("866-320-PG-01559", "A tag from a list. The serial no. and reference drawing on that row come with it."),
    ("*3CA1P01*", "Any text, such as a piping class. Searched in every line."),
    ('re:^\\d+"-RP-', "Regular expression. Start the value with re:"),
]


def sanitize_rel(name):
    name = (name or "file").replace("\\", "/")
    parts = [p for p in name.split("/") if p not in ("", ".", "..")]
    return "/".join(parts) or "file"


def clip(value, limit=400):
    text = "" if value is None else str(value)
    if len(text) <= limit:
        return text
    return text[:limit] + "..."


def job_public(job):
    return {
        "id": job["id"],
        "status": job["status"],
        "progress": job.get("progress") or {},
        "error": job.get("error") or "",
        "search_error": job.get("search_error") or "",
        "stats": job.get("stats"),
        "result": job.get("result"),
    }


def evict_old_jobs():
    finished = [jid for jid, job in JOBS.items() if job["status"] in ("ready", "error", "cancelled")]
    finished.sort(key=lambda jid: JOBS[jid]["created"])
    while len(JOBS) > MAX_JOBS and finished:
        jid = finished.pop(0)
        job = JOBS.pop(jid, None)
        if job:
            shutil.rmtree(job["dir"], ignore_errors=True)


def save_uploads(job_dir, storage_items, relative_paths):
    saved = []
    skipped = []
    in_dir = os.path.join(job_dir, "in")
    os.makedirs(in_dir, exist_ok=True)
    for index, storage in enumerate(storage_items):
        raw = relative_paths[index] if index < len(relative_paths) else storage.filename
        rel = sanitize_rel(raw or storage.filename)
        if not rel.lower().endswith((".pdf", ".zip")):
            skipped.append(rel)
            continue
        dest = os.path.join(in_dir, *rel.split("/"))
        os.makedirs(os.path.dirname(dest), exist_ok=True)
        storage.save(dest)
        saved.append(dest)
    return saved, skipped


def collect_inputs(in_dir):
    inputs = [in_dir]
    for root, _dirs, files in os.walk(in_dir):
        for name in files:
            if name.lower().endswith(".zip"):
                inputs.append(os.path.join(root, name))
    return inputs


def build_stats(datasets):
    kinds = {}
    total_pages = 0
    files = []
    for display, _path, data in datasets:
        total_pages += len(data["pages"])
        files.append({
            "name": display,
            "pages": len(data["pages"]),
            "items": len(data["items"]),
        })
        for item in data["items"]:
            kinds[item[0]] = kinds.get(item[0], 0) + 1
    return {
        "files": len(datasets),
        "pages": total_pages,
        "kinds": [{"name": name, "count": count} for name, count in sorted(kinds.items())],
        "file_list": files,
    }


def analyse_worker(job):
    try:
        job["progress"] = {
            "phase": "analyse",
            "message": "Collecting PDFs",
            "file_index": 0,
            "file_count": 0,
            "page": 0,
            "pages": 0,
            "items": 0,
            "eta": 0,
            "file_name": "",
        }
        files = eng.collect_pdfs(collect_inputs(os.path.join(job["dir"], "in")))
        if job.get("cancel"):
            raise eng.AnalysisCancelled()
        if not files:
            job["error"] = "No PDF files found. Upload a PDF, a ZIP of PDFs, or a folder of PDFs."
            job["status"] = "error"
            return

        def progress(info):
            if job.get("cancel"):
                raise eng.AnalysisCancelled()
            merged = dict(info)
            merged["phase"] = "analyse"
            job["progress"] = merged

        datasets = eng.analyze_all(files, job["ocr"], use_cache=True, progress=progress)
        if job.get("cancel"):
            raise eng.AnalysisCancelled()
        if not datasets:
            job["error"] = "None of the PDFs could be analysed."
            job["status"] = "error"
            return
        job["datasets"] = datasets
        job["stats"] = build_stats(datasets)
        job["progress"] = {
            "phase": "ready",
            "message": "Drawings are ready to search",
            "file_index": len(datasets),
            "file_count": len(datasets),
            "page": job["stats"]["pages"],
            "pages": job["stats"]["pages"],
            "items": sum(item["items"] for item in job["stats"]["file_list"]),
            "eta": 0,
            "file_name": "",
        }
        job["status"] = "ready"
    except eng.AnalysisCancelled:
        job["error"] = "Reading was cancelled."
        job["status"] = "cancelled"
    except Exception as exc:
        traceback.print_exc()
        job["error"] = str(exc) or "The drawings could not be read."
        job["status"] = "error"


def search_worker(job, queries, columns):
    try:
        job["progress"] = {
            "phase": "search",
            "message": "Matching tags and writing the workbook",
            "file_index": 0,
            "file_count": len(job["datasets"]),
            "page": 0,
            "pages": 0,
            "items": 0,
            "eta": 0,
            "file_name": "",
        }
        rows, summary, totals = eng.search_all(job["datasets"], queries)
        excel_name = ""
        download = None
        if rows:
            excel_name = "extracted_tags_%s.xlsx" % datetime.now().strftime("%Y%m%d_%H%M%S")
            excel_path = os.path.join(job["dir"], excel_name)
            eng.write_excel(
                excel_path,
                rows,
                summary,
                columns,
                "%d PDF file(s)" % len(job["datasets"]),
                job["datasets"],
            )
            job["excel"] = excel_path
            download = "/api/jobs/%s/download" % job["id"]
        else:
            job["excel"] = ""

        preview = []
        for row in rows[:PREVIEW_LIMIT]:
            preview.append({column: clip(row.get(column, "")) for column in columns})
        summary_out = []
        for item in summary:
            summary_out.append({
                "Query": item["Query"],
                "File Name": item["File Name"],
                "Matches": item["Matches"],
                "Pages": item["Pages"],
                "Status": item["Status"],
                "Page list": clip(item["Page list"], 300),
            })
        job["result"] = {
            "columns": columns,
            "rows": preview,
            "row_count": len(rows),
            "truncated": len(rows) > PREVIEW_LIMIT,
            "totals": totals,
            "summary": summary_out,
            "download": download,
            "excel_name": excel_name,
        }
        job["status"] = "ready"
    except Exception as exc:
        traceback.print_exc()
        job["search_error"] = str(exc) or "The search could not be finished."
        job["status"] = "ready"


@app.get("/")
def index():
    return render_template("index.html")


@app.get("/api/meta")
def meta():
    return jsonify({
        "columns": eng.COLUMNS,
        "default_columns": eng.DEFAULT_COLUMNS,
        "required_columns": eng.REQUIRED_COLUMNS,
        "examples": [{"value": value, "hint": hint} for value, hint in EXAMPLES],
    })


@app.post("/api/jobs")
def create_job():
    uploads = request.files.getlist("files")
    if not uploads:
        return jsonify(error="Choose at least one PDF or ZIP."), 400
    ocr = request.form.get("ocr", "auto")
    if ocr not in ("auto", "off", "all"):
        ocr = "auto"
    try:
        relative_paths = json.loads(request.form.get("paths") or "[]")
        if not isinstance(relative_paths, list):
            relative_paths = []
    except json.JSONDecodeError:
        relative_paths = []

    job_id = uuid.uuid4().hex[:12]
    job_dir = os.path.join(WORK, job_id)
    os.makedirs(job_dir, exist_ok=True)
    saved, skipped = save_uploads(job_dir, uploads, [str(p) for p in relative_paths])
    if not saved:
        shutil.rmtree(job_dir, ignore_errors=True)
        return jsonify(error="None of those files are PDFs or ZIPs."), 400

    job = {
        "id": job_id,
        "dir": job_dir,
        "created": datetime.now().timestamp(),
        "status": "analyzing",
        "ocr": ocr,
        "cancel": False,
        "progress": {"phase": "analyse", "message": "Starting", "page": 0, "pages": 0},
        "error": "",
        "search_error": "",
        "stats": None,
        "result": None,
        "datasets": None,
        "excel": "",
        "skipped": skipped,
    }
    with LOCK:
        JOBS[job_id] = job
        evict_old_jobs()
    threading.Thread(target=analyse_worker, args=(job,), daemon=True).start()
    return jsonify(id=job_id, skipped=skipped)


@app.get("/api/jobs/<job_id>")
def job_status(job_id):
    job = JOBS.get(job_id)
    if not job:
        return jsonify(error="This session expired. Choose the files again."), 404
    return jsonify(job_public(job))


@app.post("/api/jobs/<job_id>/cancel")
def cancel_job(job_id):
    job = JOBS.get(job_id)
    if not job:
        return jsonify(ok=True)
    job["cancel"] = True
    return jsonify(ok=True)


@app.post("/api/jobs/<job_id>/search")
def start_search(job_id):
    job = JOBS.get(job_id)
    if not job:
        return jsonify(error="This session expired. Choose the files again."), 404
    if job["status"] == "analyzing":
        return jsonify(error="The drawings are still being read."), 409
    if job["status"] == "searching":
        return jsonify(error="A search is already running."), 409
    if job["status"] != "ready" or not job.get("datasets"):
        return jsonify(error="Read the drawings before searching."), 409

    payload = request.get_json(silent=True) or {}
    queries = eng.split_csv(payload.get("tags") or "")
    if not queries:
        return jsonify(error="Enter at least one tag."), 400
    bad = []
    for query in queries:
        try:
            eng.build_matcher(query)
        except re.error as exc:
            bad.append("%s (%s)" % (query, exc))
    if bad:
        return jsonify(error="Bad regular expression: " + "; ".join(bad)), 400
    columns, unknown = eng.resolve_columns(payload.get("columns") or [])
    if unknown:
        return jsonify(error="Unknown column: " + ", ".join(unknown)), 400

    job["result"] = None
    job["search_error"] = ""
    job["status"] = "searching"
    threading.Thread(target=search_worker, args=(job, queries, columns), daemon=True).start()
    return jsonify(ok=True)


@app.get("/api/jobs/<job_id>/download")
def download_excel(job_id):
    job = JOBS.get(job_id)
    path = job.get("excel") if job else ""
    if not path or not os.path.isfile(path):
        return jsonify(error="There is no workbook to download yet."), 404
    return send_file(path, as_attachment=True, download_name=os.path.basename(path))


@app.errorhandler(413)
def too_large(_exc):
    return jsonify(error="That upload is larger than 2 GB."), 413


def pick_port(start=5050):
    for port in range(start, start + 20):
        with socket.socket() as sock:
            try:
                sock.bind(("127.0.0.1", port))
            except OSError:
                continue
            return port
    return start


def serve():
    os.makedirs(WORK, exist_ok=True)
    port = pick_port()
    url = "http://127.0.0.1:%d" % port
    print("=" * 66)
    print("   PDF / P&ID TAG EXTRACTOR")
    print("   Open this address in your browser:")
    print("   %s" % url)
    print("   Press Ctrl+C to stop.")
    print("=" * 66)
    threading.Timer(0.8, lambda: webbrowser.open(url)).start()
    try:
        from waitress import serve as waitress_serve
        waitress_serve(app, host="127.0.0.1", port=port, threads=8, channel_timeout=7200)
    except ImportError:
        app.run(host="127.0.0.1", port=port, threaded=True, use_reloader=False)


if __name__ == "__main__":
    serve()
