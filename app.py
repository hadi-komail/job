import json
import os
import subprocess
import sys
from datetime import datetime
from io import BytesIO
from pathlib import Path

from flask import Flask, jsonify, redirect, render_template_string, request, send_file, send_from_directory, url_for
from supabase import create_client
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Pt, RGBColor

app = Flask(__name__)
BASE_DIR = Path(__file__).resolve().parent
LETTERS_DIR = BASE_DIR / "letters"
JOB_DESCRIPTIONS_DIR = BASE_DIR / "job_descriptions"
TEMPLATE_PATH = BASE_DIR / "cover_letter_template.docx"
LOG_FILE = BASE_DIR / "run.log"
META_FILE = BASE_DIR / "job_meta.json"
RUN_STATE_FILE = BASE_DIR / "run_state.json"
SUPABASE_TABLE = "jobs"
STATUS_ORDER = ["not_applied", "to_apply", "applied", "not_to_apply"]
STATUS_LABELS = {
    "not_applied": "Not Applied",
    "to_apply": "To Apply",
    "applied": "Applied",
    "not_to_apply": "Not to Apply",
}
LETTER_FONT_NAME = "Helvetica"
LETTER_FONT_SIZE = 9
LETTER_SPACE_AFTER_PT = 6
LETTER_TEXT_COLOR = RGBColor(29, 39, 49)
LETTER_LINE_SPACING = 1.5

current_process = None


def is_running():
    global current_process
    refresh_run_state()
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


def load_run_state():
    return load_json(
        RUN_STATE_FILE,
        {
            "status": "idle",
            "started_at": "",
            "finished_at": "",
            "returncode": None,
        },
    )


def save_run_state(status, *, returncode=None):
    state = load_run_state()
    state["status"] = status
    if status == "running":
        state["started_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        state["finished_at"] = ""
        state["returncode"] = None
    else:
        state["finished_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        state["returncode"] = returncode
    save_json(RUN_STATE_FILE, state)


def refresh_run_state():
    global current_process
    if current_process is None:
        return
    returncode = current_process.poll()
    if returncode is None:
        return
    save_run_state("completed" if returncode == 0 else "failed", returncode=returncode)
    current_process = None


def get_supabase_client():
    url = os.environ.get("SUPABASE_URL")
    key = os.environ.get("SUPABASE_ANON_KEY")
    if not url or not key:
        return None
    return create_client(url, key)


def normalize_path_string(path_string):
    if not path_string:
        return None
    relative = Path(str(path_string).replace("\\", "/"))
    return BASE_DIR / relative


def rel_download_path(path):
    return str(path.relative_to(BASE_DIR)).replace("\\", "/")


def load_jobs():
    client = get_supabase_client()
    if client is None:
        return []
    response = client.table(SUPABASE_TABLE).select("*").order("date", desc=True).execute()
    rows = response.data or []
    jobs = []
    for row in rows:
        refnr = str(row.get("refnr", "")).strip()
        if not refnr:
            continue
        cover_letter_path = normalize_path_string(row.get("cover_letter_path"))
        job_description_path = normalize_path_string(row.get("job_description_path"))
        jobs.append(
            {
                "refnr": refnr,
                "job_id": row.get("job_id") or "",
                "date": row.get("date") or "",
                "keyword_score": row.get("keyword_score") or "",
                "ai_match_score": row.get("ai_match_score") or "",
                "title": row.get("title") or "",
                "employer": row.get("employer") or "",
                "city": row.get("city") or "",
                "reason": row.get("reason") or "",
                "job_url": row.get("job_url") or "",
                "cover_letter_path": cover_letter_path,
                "job_description_path": job_description_path,
                "cover_letter_text": row.get("cover_letter_text") or "",
                "job_description_text": row.get("job_description_text") or "",
                "has_cover_letter": bool(row.get("has_cover_letter")),
                "application_status": row.get("application_status", "not_applied"),
                "application_status_label": STATUS_LABELS.get(
                    row.get("application_status", "not_applied"),
                    "Not Applied",
                ),
                "application_result": row.get("application_result", ""),
                "note": row.get("note", ""),
                "updated_at": row.get("updated_at", ""),
            }
        )

    def sort_key(item):
        return (
            0 if item["has_cover_letter"] else 1,
            str(item["job_id"] or ""),
            str(item["date"] or ""),
        )

    jobs.sort(key=sort_key)
    return jobs


def fetch_job(refnr):
    client = get_supabase_client()
    if client is None:
        return None
    response = client.table(SUPABASE_TABLE).select("*").eq("refnr", str(refnr)).limit(1).execute()
    rows = response.data or []
    return rows[0] if rows else None


def german_date_string():
    months = [
        "Januar",
        "Februar",
        "Maerz",
        "April",
        "Mai",
        "Juni",
        "Juli",
        "August",
        "September",
        "Oktober",
        "November",
        "Dezember",
    ]
    today = datetime.now()
    return f"{today.day}. {months[today.month - 1]} {today.year}"


def style_paragraph(paragraph, *, space_after_pt=LETTER_SPACE_AFTER_PT):
    paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(space_after_pt)
    paragraph.paragraph_format.line_spacing = LETTER_LINE_SPACING
    for run in paragraph.runs:
        run.font.name = LETTER_FONT_NAME
        run.font.size = Pt(LETTER_FONT_SIZE)
        run.font.color.rgb = LETTER_TEXT_COLOR


def style_table_paragraph(paragraph):
    paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.line_spacing = LETTER_LINE_SPACING
    for run in paragraph.runs:
        run.font.name = LETTER_FONT_NAME
        run.font.size = Pt(LETTER_FONT_SIZE)
        run.font.color.rgb = LETTER_TEXT_COLOR


def replace_paragraph_text(paragraph, text, *, bold=None, space_after_pt=LETTER_SPACE_AFTER_PT):
    paragraph.text = ""
    run = paragraph.add_run(text)
    if bold is not None:
        run.bold = bold
    style_paragraph(paragraph, space_after_pt=space_after_pt)


def insert_paragraph_after(paragraph, text):
    new_paragraph = paragraph.insert_paragraph_before("")
    paragraph._p.addnext(new_paragraph._p)
    new_paragraph.style = paragraph.style
    if paragraph.alignment is not None:
        new_paragraph.alignment = paragraph.alignment
    new_paragraph.add_run(text)
    style_paragraph(new_paragraph)
    return new_paragraph


def restyle_document(document):
    for paragraph in document.paragraphs:
        if paragraph.text.strip():
            style_paragraph(paragraph)

    for table in document.tables:
        for row_cells in table.rows:
            row_cells.height = None
            row_cells.height_rule = None
            for cell in row_cells.cells:
                for paragraph in cell.paragraphs:
                    style_table_paragraph(paragraph)


def fill_cover_letter_template(document, row):
    subject = f"Bewerbung als {row.get('title') or 'Stelle'}"
    body_text = (row.get("cover_letter_text") or "").strip()
    body_parts = [part.strip() for part in body_text.split("\n\n") if part.strip()] or [body_text]

    employer = row.get("employer") or ""
    city = row.get("city") or ""

    replacements = {
        "{{DATE}}": german_date_string(),
        "{{SUBJECT}}": subject,
        "{{EMPLOYER}}": employer,
        "{{STREET, HOUSE NUMBER}}": "",
        "{{POSTAL NUMBER, CITY}}": city,
    }

    for paragraph in document.paragraphs:
        text = (paragraph.text or "").strip()
        if text in replacements:
            replace_paragraph_text(
                paragraph,
                replacements[text],
                bold=(text == "{{SUBJECT}}"),
            )
        elif text == "{{BODY}}":
            replace_paragraph_text(paragraph, body_parts[0] if body_parts else "")
            current_paragraph = paragraph
            for extra in body_parts[1:]:
                current_paragraph = insert_paragraph_after(current_paragraph, extra)

    for table in document.tables:
        for row_cells in table.rows:
            row_cells.height = None
            row_cells.height_rule = None
            for cell in row_cells.cells:
                for paragraph in cell.paragraphs:
                    text = (paragraph.text or "").strip()
                    if text in replacements:
                        replace_paragraph_text(paragraph, replacements[text], space_after_pt=0)
                        style_table_paragraph(paragraph)
                    else:
                        style_table_paragraph(paragraph)

    restyle_document(document)


def group_jobs_by_status(jobs):
    grouped = {status: [] for status in STATUS_ORDER}
    for job in jobs:
        status = job.get("application_status") or "not_applied"
        if status not in grouped:
            status = "not_applied"
        grouped[status].append(job)

    return [
        {
            "key": status,
            "label": STATUS_LABELS[status],
            "jobs": grouped[status],
        }
        for status in STATUS_ORDER
    ]


@app.get("/")
def dashboard():
    jobs = load_jobs()
    grouped_sections = group_jobs_by_status(jobs)
    application_results = [job for job in jobs if (job.get("application_result") or "").strip()]
    running = is_running()
    run_state = load_run_state()
    status = "Running" if running else run_state.get("status", "idle").title()
    logs = LOG_FILE.read_text(encoding="utf-8", errors="replace") if LOG_FILE.exists() else "No log file yet."
    recent_logs = "\n".join(logs.splitlines()[-40:])

    template = """
    <!doctype html>
    <html lang="en">
    <head>
      <meta charset="utf-8">
      <meta name="viewport" content="width=device-width, initial-scale=1">
      <title>Job Automation</title>
      {% if status == "Running" %}
      <meta http-equiv="refresh" content="5">
      {% endif %}
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
          min-height: 180px;
          max-height: 420px;
          overflow: auto;
          background: #0e2133;
          color: #d7e7f5;
          border-radius: 14px;
          padding: 14px;
        }
        .log-toolbar {
          display: flex;
          flex-wrap: wrap;
          gap: 10px;
          margin-bottom: 12px;
        }
        .inline-note {
          font-size: 12px;
          color: var(--muted);
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
            <div class="tiny">Started: {{ run_state.started_at or "—" }}</div>
            <div class="tiny" style="margin-bottom: 10px;">Finished: {{ run_state.finished_at or "—" }}</div>
            <a class="button" href="/run">{{ "Already Running" if status == "Running" else "Run Job Search" }}</a>
          </div>
          <div class="panel">
            <div class="tiny">Jobs in dashboard</div>
            <div style="font-size: 30px; font-weight: 800; margin-top: 6px;">{{ jobs|length }}</div>
            <div class="tiny">Loaded from Supabase</div>
          </div>
          <div class="panel">
            <div class="tiny">Cover letters</div>
            <div style="font-size: 30px; font-weight: 800; margin-top: 6px;">{{ jobs|selectattr("has_cover_letter")|list|length }}</div>
            <div class="tiny">Jobs with generated letters</div>
          </div>
        </div>

        <div class="section-title">Recent Logs</div>
        <div class="panel">
          <div class="log-toolbar">
            <a class="button secondary" style="width:auto;" href="/logs" target="_blank" rel="noopener">Open Full Logs</a>
            <div class="inline-note">
              {% if status == "Running" %}
              Auto-refreshing every 5 seconds while running.
              {% else %}
              Refresh the page after a new run to see updated logs.
              {% endif %}
            </div>
          </div>
          <div class="logs">{{ recent_logs }}</div>
        </div>

        {% for section in grouped_sections %}
        <div class="section-title">{{ section.label }}</div>
        <div class="jobs">
          {% for job in section.jobs %}
          <div class="job">
            <div class="job-top">
              <div class="job-title">{{ job.title }}</div>
              <div class="meta">
                <span class="chip">{{ job.job_id or job.refnr }}</span>
                <span class="chip">{{ job.employer }}</span>
                <span class="chip">{{ job.city }}</span>
                <span class="chip">Date: {{ job.date }}</span>
                <span class="chip">Keyword score: {{ job.keyword_score }}</span>
                <span class="chip">AI match: {{ job.ai_match_score }}/10</span>
                <span class="chip">{{ job.application_status_label }}</span>
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

                  {% if job.has_cover_letter %}
                  <a href="{{ url_for('download_cover_letter', refnr=job.refnr) }}">Download cover letter</a>
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
                      <option value="to_apply" {{ "selected" if job.application_status == "to_apply" else "" }}>To apply</option>
                      <option value="applied" {{ "selected" if job.application_status == "applied" else "" }}>Applied</option>
                      <option value="not_to_apply" {{ "selected" if job.application_status == "not_to_apply" else "" }}>Not to apply</option>
                    </select>

                    <label class="tiny" for="result-{{ job.refnr }}">Application result</label>
                    <textarea id="result-{{ job.refnr }}" name="application_result" placeholder="Interview, rejection, no reply, accepted...">{{ job.application_result }}</textarea>

                    <label class="tiny" for="note-{{ job.refnr }}">Note</label>
                    <textarea id="note-{{ job.refnr }}" name="note" placeholder="Add your note here...">{{ job.note }}</textarea>

                    <button class="button secondary" type="submit">Save</button>
                    <div class="tiny">Job ID: {{ job.job_id or "Pending" }}</div>
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
          <div class="panel">No jobs in this section yet.</div>
          {% endfor %}
        </div>
        {% endfor %}

        <div class="section-title">Application Results</div>
        <div class="jobs">
          {% for job in application_results %}
          <div class="job">
            <div class="job-top">
              <div class="job-title">{{ job.title }}</div>
              <div class="meta">
                <span class="chip">{{ job.job_id or job.refnr }}</span>
                <span class="chip">{{ job.employer }}</span>
                <span class="chip">{{ job.application_status_label }}</span>
              </div>
            </div>
            <div class="job-body">
              <div class="stack">
                <div class="card">
                  <h3>Application Outcome</h3>
                  <p class="reason">{{ job.application_result }}</p>
                </div>
              </div>
              <div class="stack">
                <div class="card">
                  <h3>Quick Links</h3>
                  <div class="links">
                    {% if job.job_url %}
                    <a href="{{ job.job_url }}" target="_blank" rel="noopener">Open job listing</a>
                    {% endif %}
                    {% if job.cover_letter_path %}
                    <a href="{{ url_for('download_cover_letter', refnr=job.refnr) }}">Download cover letter</a>
                    {% endif %}
                  </div>
                </div>
              </div>
            </div>
          </div>
          {% else %}
          <div class="panel">No application outcomes recorded yet.</div>
          {% endfor %}
        </div>
      </div>
    </body>
    </html>
    """
    return render_template_string(
        template,
        jobs=jobs,
        grouped_sections=grouped_sections,
        application_results=application_results,
        status=status,
        recent_logs=recent_logs,
        run_state=run_state,
    )


@app.get("/run")
def run_script():
    global current_process

    if is_running():
        return redirect(url_for("dashboard"))

    log_handle = open(LOG_FILE, "w", encoding="utf-8")
    env = os.environ.copy()
    env["PYTHONUNBUFFERED"] = "1"
    current_process = subprocess.Popen(
        [sys.executable, "-u", str(BASE_DIR / "main.py")],
        cwd=BASE_DIR,
        stdout=log_handle,
        stderr=subprocess.STDOUT,
        text=True,
        bufsize=1,
        env=env,
    )
    save_run_state("running")
    return redirect(url_for("dashboard"))


@app.get("/logs")
def logs():
    running = is_running()
    run_state = load_run_state()
    content = LOG_FILE.read_text(encoding="utf-8", errors="replace") if LOG_FILE.exists() else "No log file yet."
    return render_template_string(
        """
        <!doctype html>
        <html lang="en">
        <head>
          <meta charset="utf-8">
          <meta name="viewport" content="width=device-width, initial-scale=1">
          <title>Run Logs</title>
          {% if running %}
          <meta http-equiv="refresh" content="5">
          {% endif %}
          <style>
            body {
              margin: 0;
              padding: 18px;
              background: #0c1b2a;
              color: #d7e7f5;
              font-family: Consolas, monospace;
            }
            .top {
              margin-bottom: 14px;
              font-family: "Segoe UI", Arial, sans-serif;
            }
            a {
              color: #8dc5ff;
            }
            pre {
              white-space: pre-wrap;
              line-height: 1.45;
              font-size: 13px;
            }
          </style>
        </head>
        <body>
          <div class="top">
            <div><strong>Status:</strong> {{ run_state.status }}</div>
            <div><strong>Started:</strong> {{ run_state.started_at or "—" }}</div>
            <div><strong>Finished:</strong> {{ run_state.finished_at or "—" }}</div>
            <div><a href="/">Back to dashboard</a></div>
          </div>
          <pre>{{ content }}</pre>
        </body>
        </html>
        """,
        content=content,
        running=running,
        run_state=run_state,
    )


@app.post("/jobs/<refnr>/update")
def update_job(refnr):
    client = get_supabase_client()
    if client is not None:
        client.table(SUPABASE_TABLE).update(
            {
                "application_status": request.form.get("application_status", "not_applied"),
                "application_result": request.form.get("application_result", "").strip(),
                "note": request.form.get("note", "").strip(),
                "updated_at": datetime.now().isoformat(),
            }
        ).eq("refnr", str(refnr)).execute()
    return redirect(url_for("dashboard"))


@app.patch("/api/jobs/<refnr>")
def update_job_api(refnr):
    client = get_supabase_client()
    if client is None:
        return jsonify({"error": "Supabase is not configured"}), 500

    payload = request.get_json(silent=True) or {}
    application_status = payload.get("application_status", "not_applied")
    application_result = str(payload.get("application_result", "")).strip()
    note = str(payload.get("note", "")).strip()
    updated_at = payload.get("updated_at") or datetime.now().isoformat()

    response = client.table(SUPABASE_TABLE).update(
        {
            "application_status": application_status,
            "application_result": application_result,
            "note": note,
            "updated_at": updated_at,
        }
    ).eq("refnr", str(refnr)).execute()
    return jsonify({"updated": True, "data": response.data or []})


@app.get("/download/cover-letter/<refnr>")
def download_cover_letter(refnr):
    row = fetch_job(refnr)
    if not row:
        return "Cover letter not found", 404

    cover_letter_text = (row.get("cover_letter_text") or "").strip()
    if not cover_letter_text:
        return "Cover letter text not found", 404
    if not TEMPLATE_PATH.exists():
        return "Cover letter template not found", 500

    title = row.get("title") or "job"
    employer = row.get("employer") or "employer"
    file_name = f"{refnr}_{employer}_{title}.docx"
    safe_name = "".join(ch if ch.isalnum() or ch in ("-", "_", ".") else "_" for ch in file_name)

    document = Document(TEMPLATE_PATH)
    fill_cover_letter_template(document, row)

    buffer = BytesIO()
    document.save(buffer)
    buffer.seek(0)
    return send_file(
        buffer,
        as_attachment=True,
        download_name=safe_name,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


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
