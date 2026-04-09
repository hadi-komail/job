import json
import subprocess
import sys
from datetime import datetime
from pathlib import Path

from flask import Flask, redirect, render_template_string, request, send_from_directory, url_for
from openpyxl import load_workbook

app = Flask(__name__)
BASE_DIR = Path(__file__).resolve().parent
LETTERS_DIR = BASE_DIR / "letters"
JOB_DESCRIPTIONS_DIR = BASE_DIR / "job_descriptions"
LOG_FILE = BASE_DIR / "run.log"
META_FILE = BASE_DIR / "job_meta.json"
AI_SCORED_JOBS_PATH = BASE_DIR / "ai_scored_jobs.xlsx"
AI_WRITTEN_JOBS_PATH = BASE_DIR / "ai_cover_letters.xlsx"

current_process = None


def is_running():
    global current_process
    return current_process is not None and current_process.poll() is None


def load_json(path, default):
    if not path.exists():
        return default
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return default


def save_json(path, data):
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def read_workbook_rows(path):
    if not path.exists():
        return []
    workbook = load_workbook(path)
    sheet = workbook.active
    rows = list(sheet.iter_rows(values_only=True))
    if not rows:
        return []
    headers = list(rows[0])
    items = []
    for row in rows[1:]:
        items.append({headers[i]: row[i] for i in range(len(headers))})
    return items


def normalize_path_string(path_string):
    if not path_string:
        return None
    relative = Path(str(path_string).replace("\\", "/"))
    return BASE_DIR / relative


def rel_download_path(path):
    return str(path.relative_to(BASE_DIR)).replace("\\", "/")


def load_jobs():
    scored_rows = read_workbook_rows(AI_SCORED_JOBS_PATH)
    written_rows = read_workbook_rows(AI_WRITTEN_JOBS_PATH)
    meta = load_json(META_FILE, {})

    written_by_refnr = {str(row["refnr"]): row for row in written_rows if row.get("refnr")}
    jobs = []

    for row in scored_rows:
        refnr = str(row.get("refnr", "")).strip()
        if not refnr:
            continue

        written = written_by_refnr.get(refnr, {})
        cover_letter_path = normalize_path_string(written.get("cover_letter_path"))
        job_description_path = None
        if cover_letter_path:
            candidate = JOB_DESCRIPTIONS_DIR / cover_letter_path.name
            if candidate.exists():
                job_description_path = candidate

        meta_item = meta.get(refnr, {})
        jobs.append(
            {
                "refnr": refnr,
                "date": row.get("date") or "",
                "keyword_score": row.get("keyword_score") or "",
                "ai_match_score": row.get("ai_match_score") or "",
                "title": row.get("title") or "",
                "employer": row.get("employer") or "",
                "city": row.get("city") or "",
                "reason": row.get("reason") or "",
                "job_url": written.get("job_url") or "",
                "cover_letter_path": cover_letter_path,
                "job_description_path": job_description_path,
                "has_cover_letter": bool(written),
                "application_status": meta_item.get("application_status", "not_applied"),
                "note": meta_item.get("note", ""),
                "updated_at": meta_item.get("updated_at", ""),
            }
        )

    def sort_key(item):
        return (
            0 if item["has_cover_letter"] else 1,
            str(item["date"] or ""),
        )

    jobs.sort(key=sort_key, reverse=True)
    return jobs


@app.get("/")
def dashboard():
    jobs = load_jobs()
    status = "Running" if is_running() else "Idle"
    logs = LOG_FILE.read_text(encoding="utf-8", errors="replace") if LOG_FILE.exists() else "No log file yet."
    recent_logs = "\n".join(logs.splitlines()[-40:])

    template = """
    <!doctype html>
    <html lang="en">
    <head>
      <meta charset="utf-8">
      <meta name="viewport" content="width=device-width, initial-scale=1">
      <title>Job Automation</title>
      <style>
        :root {
          --bg: #eef3f8;
          --panel: #ffffff;
          --ink: #18324b;
          --muted: #5b7288;
          --line: #d7e1ea;
          --accent: #0f5ea8;
          --accent-2: #dbeafe;
          --success: #1f7a4d;
          --warn: #9a6700;
          --shadow: 0 18px 50px rgba(17, 36, 56, 0.10);
          --radius: 18px;
        }
        * { box-sizing: border-box; }
        body {
          margin: 0;
          font-family: "Segoe UI", Arial, sans-serif;
          background:
            radial-gradient(circle at top left, #f8fbff 0, #eef3f8 48%, #e8eef5 100%);
          color: var(--ink);
        }
        .wrap {
          max-width: 1180px;
          margin: 0 auto;
          padding: 20px 14px 40px;
        }
        .hero {
          background: linear-gradient(135deg, #113350, #245f95);
          color: white;
          border-radius: 24px;
          padding: 22px 20px;
          box-shadow: var(--shadow);
          margin-bottom: 18px;
        }
        .hero h1 {
          margin: 0 0 10px;
          font-size: 28px;
        }
        .hero p {
          margin: 0;
          color: rgba(255,255,255,0.86);
        }
        .toolbar {
          display: grid;
          grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
          gap: 14px;
          margin: 18px 0;
        }
        .panel {
          background: var(--panel);
          border: 1px solid var(--line);
          border-radius: var(--radius);
          box-shadow: var(--shadow);
          padding: 16px;
        }
        .status-badge {
          display: inline-block;
          padding: 8px 12px;
          border-radius: 999px;
          background: {{ '#dff7ea' if status == 'Idle' else '#fff3cd' }};
          color: {{ '#1f7a4d' if status == 'Idle' else '#9a6700' }};
          font-weight: 700;
        }
        .button {
          display: inline-flex;
          align-items: center;
          justify-content: center;
          width: 100%;
          padding: 12px 14px;
          border-radius: 14px;
          text-decoration: none;
          font-weight: 700;
          border: none;
          cursor: pointer;
          background: var(--accent);
          color: white;
        }
        .button.secondary {
          background: #edf4fb;
          color: var(--accent);
        }
        .button:disabled {
          opacity: 0.6;
          cursor: not-allowed;
        }
        .section-title {
          margin: 18px 0 12px;
          font-size: 18px;
          font-weight: 800;
        }
        .logs {
          white-space: pre-wrap;
          font-family: Consolas, monospace;
          font-size: 12px;
          line-height: 1.4;
          max-height: 280px;
          overflow: auto;
          background: #0e2133;
          color: #d7e7f5;
          border-radius: 14px;
          padding: 14px;
        }
        .jobs {
          display: grid;
          gap: 16px;
        }
        .job {
          background: var(--panel);
          border: 1px solid var(--line);
          border-radius: 20px;
          box-shadow: var(--shadow);
          overflow: hidden;
        }
        .job-top {
          padding: 18px 18px 12px;
          border-bottom: 1px solid var(--line);
          background: linear-gradient(180deg, #fbfdff, #f3f7fb);
        }
        .job-title {
          font-size: 20px;
          font-weight: 800;
          margin: 0 0 8px;
        }
        .meta {
          display: flex;
          flex-wrap: wrap;
          gap: 8px;
        }
        .chip {
          background: #eaf2fa;
          color: #23486b;
          border-radius: 999px;
          padding: 6px 10px;
          font-size: 12px;
          font-weight: 700;
        }
        .job-body {
          display: grid;
          grid-template-columns: 1.2fr 0.8fr;
          gap: 16px;
          padding: 16px 18px 18px;
        }
        .stack {
          display: grid;
          gap: 12px;
        }
        .card {
          border: 1px solid var(--line);
          border-radius: 14px;
          padding: 14px;
          background: #fff;
        }
        .card h3 {
          margin: 0 0 10px;
          font-size: 14px;
          text-transform: uppercase;
          letter-spacing: 0.04em;
          color: var(--muted);
        }
        .links a {
          display: block;
          padding: 8px 0;
          color: var(--accent);
          text-decoration: none;
          word-break: break-word;
        }
        .reason {
          margin: 0;
          line-height: 1.55;
          color: var(--ink);
        }
        form {
          display: grid;
          gap: 10px;
        }
        select, textarea {
          width: 100%;
          border: 1px solid var(--line);
          border-radius: 12px;
          padding: 10px 12px;
          font: inherit;
          color: var(--ink);
          background: #fff;
        }
        textarea {
          min-height: 110px;
          resize: vertical;
        }
        .tiny {
          font-size: 12px;
          color: var(--muted);
        }
        @media (max-width: 900px) {
          .job-body {
            grid-template-columns: 1fr;
          }
        }
      </style>
    </head>
    <body>
      <div class="wrap">
        <div class="hero">
          <h1>Job Automation Dashboard</h1>
          <p>Run the search, review AI-evaluated jobs, download files, and track your application status from your phone.</p>
        </div>

        <div class="toolbar">
          <div class="panel">
            <div class="tiny">Runner</div>
            <div style="margin: 8px 0 14px;"><span class="status-badge">{{ status }}</span></div>
            <a class="button" href="/run">{{ "Already Running" if status == "Running" else "Run Job Search" }}</a>
          </div>
          <div class="panel">
            <div class="tiny">Jobs in dashboard</div>
            <div style="font-size: 30px; font-weight: 800; margin-top: 6px;">{{ jobs|length }}</div>
            <div class="tiny">Loaded from Excel tracking files</div>
          </div>
          <div class="panel">
            <div class="tiny">Cover letters</div>
            <div style="font-size: 30px; font-weight: 800; margin-top: 6px;">{{ jobs|selectattr("has_cover_letter")|list|length }}</div>
            <div class="tiny">Jobs with generated letters</div>
          </div>
        </div>

        <div class="section-title">Recent Logs</div>
        <div class="panel">
          <div class="logs">{{ recent_logs }}</div>
        </div>

        <div class="section-title">Tracked Jobs</div>
        <div class="jobs">
          {% for job in jobs %}
          <div class="job">
            <div class="job-top">
              <div class="job-title">{{ job.title }}</div>
              <div class="meta">
                <span class="chip">{{ job.employer }}</span>
                <span class="chip">{{ job.city }}</span>
                <span class="chip">Date: {{ job.date }}</span>
                <span class="chip">Keyword score: {{ job.keyword_score }}</span>
                <span class="chip">AI match: {{ job.ai_match_score }}/10</span>
                <span class="chip">{{ "Cover letter ready" if job.has_cover_letter else "No cover letter" }}</span>
              </div>
            </div>
            <div class="job-body">
              <div class="stack">
                <div class="card links">
                  <h3>Links</h3>
                  {% if job.job_url %}
                  <a href="{{ job.job_url }}" target="_blank" rel="noopener">Open job listing</a>
                  {% else %}
                  <div class="tiny">No job URL recorded yet.</div>
                  {% endif %}

                  {% if job.job_description_path %}
                  <a href="{{ url_for('download_file', folder='job_descriptions', filename=job.job_description_path.name) }}">Download job description</a>
                  {% else %}
                  <div class="tiny">No saved job description.</div>
                  {% endif %}

                  {% if job.cover_letter_path %}
                  <a href="{{ url_for('download_file', folder='letters', filename=job.cover_letter_path.name) }}">Download cover letter</a>
                  {% else %}
                  <div class="tiny">No cover letter file.</div>
                  {% endif %}
                </div>

                <div class="card">
                  <h3>AI Reason</h3>
                  <p class="reason">{{ job.reason }}</p>
                </div>
              </div>

              <div class="stack">
                <div class="card">
                  <h3>Application Tracker</h3>
                  <form method="post" action="{{ url_for('update_job', refnr=job.refnr) }}">
                    <label class="tiny" for="status-{{ job.refnr }}">Application status</label>
                    <select id="status-{{ job.refnr }}" name="application_status">
                      <option value="not_applied" {{ "selected" if job.application_status == "not_applied" else "" }}>Not applied</option>
                      <option value="applied" {{ "selected" if job.application_status == "applied" else "" }}>Applied</option>
                    </select>

                    <label class="tiny" for="note-{{ job.refnr }}">Note</label>
                    <textarea id="note-{{ job.refnr }}" name="note" placeholder="Add your note here...">{{ job.note }}</textarea>

                    <button class="button secondary" type="submit">Save</button>
                    <div class="tiny">Refnr: {{ job.refnr }}</div>
                    {% if job.updated_at %}
                    <div class="tiny">Last updated: {{ job.updated_at }}</div>
                    {% endif %}
                  </form>
                </div>
              </div>
            </div>
          </div>
          {% else %}
          <div class="panel">No jobs found yet. Run the job search first.</div>
          {% endfor %}
        </div>
      </div>
    </body>
    </html>
    """
    return render_template_string(template, jobs=jobs, status=status, recent_logs=recent_logs)


@app.get("/run")
def run_script():
    global current_process

    if is_running():
        return redirect(url_for("dashboard"))

    log_handle = open(LOG_FILE, "w", encoding="utf-8")
    current_process = subprocess.Popen(
        [sys.executable, str(BASE_DIR / "main.py")],
        cwd=BASE_DIR,
        stdout=log_handle,
        stderr=subprocess.STDOUT,
        text=True,
    )
    return redirect(url_for("dashboard"))


@app.post("/jobs/<refnr>/update")
def update_job(refnr):
    meta = load_json(META_FILE, {})
    meta[str(refnr)] = {
        "application_status": request.form.get("application_status", "not_applied"),
        "note": request.form.get("note", "").strip(),
        "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }
    save_json(META_FILE, meta)
    return redirect(url_for("dashboard"))


@app.get("/download/<folder>/<path:filename>")
def download_file(folder, filename):
    allowed = {
        "letters": LETTERS_DIR,
        "job_descriptions": JOB_DESCRIPTIONS_DIR,
    }
    if folder not in allowed:
        return "Folder not allowed", 404
    return send_from_directory(allowed[folder], filename, as_attachment=True)


if __name__ == "__main__":
    app.run(debug=True)
