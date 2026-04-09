from flask import Flask, send_from_directory
import subprocess
import sys
from pathlib import Path

app = Flask(__name__)
BASE_DIR = Path(__file__).resolve().parent
LETTERS_DIR = BASE_DIR / "letters"
JOB_DESCRIPTIONS_DIR = BASE_DIR / "job_descriptions"
LOG_FILE = BASE_DIR / "run.log"

current_process = None


def is_running():
    global current_process
    return current_process is not None and current_process.poll() is None


@app.get("/")
def home():
    status = "Running" if is_running() else "Idle"
    return f"""
    <h1>Job Automation</h1>
    <p>Status: <strong>{status}</strong></p>
    <p><a href="/run">Run job search</a></p>
    <p><a href="/status">Check status</a></p>
    <p><a href="/logs">View logs</a></p>
    <p><a href="/files">View files</a></p>
    """


@app.get("/run")
def run_script():
    global current_process

    if is_running():
        return """
        <h2>Job search is already running.</h2>
        <p><a href="/status">Check status</a></p>
        <p><a href="/logs">View logs</a></p>
        """

    log_handle = open(LOG_FILE, "w", encoding="utf-8")

    current_process = subprocess.Popen(
        [sys.executable, str(BASE_DIR / "main.py")],
        cwd=BASE_DIR,
        stdout=log_handle,
        stderr=subprocess.STDOUT,
        text=True,
    )

    return """
    <h2>Job search started.</h2>
    <p><a href="/status">Check status</a></p>
    <p><a href="/logs">View logs</a></p>
    <p><a href="/files">View files</a></p>
    """


@app.get("/status")
def status():
    if is_running():
        return """
        <h2>Status: Running</h2>
        <p><a href="/logs">View logs</a></p>
        <p><a href="/files">View files</a></p>
        """
    return """
    <h2>Status: Idle</h2>
    <p><a href="/run">Run job search</a></p>
    <p><a href="/logs">View logs</a></p>
    <p><a href="/files">View files</a></p>
    """


@app.get("/logs")
def logs():
    if LOG_FILE.exists():
        content = LOG_FILE.read_text(encoding="utf-8", errors="replace")
    else:
        content = "No log file yet."
    return f"<pre>{content}</pre>"


@app.get("/files")
def files():
    parts = ["<h1>Generated Files</h1>"]

    parts.append("<h2>Cover Letters</h2>")
    if LETTERS_DIR.exists():
        letter_files = sorted(LETTERS_DIR.iterdir(), key=lambda p: p.name.lower())
        if letter_files:
            for f in letter_files:
                parts.append(f'<p><a href="/download/letters/{f.name}">{f.name}</a></p>')
        else:
            parts.append("<p>No cover letters yet.</p>")
    else:
        parts.append("<p>No cover letters folder yet.</p>")

    parts.append("<h2>Job Descriptions</h2>")
    if JOB_DESCRIPTIONS_DIR.exists():
        desc_files = sorted(JOB_DESCRIPTIONS_DIR.iterdir(), key=lambda p: p.name.lower())
        if desc_files:
            for f in desc_files:
                parts.append(f'<p><a href="/download/job_descriptions/{f.name}">{f.name}</a></p>')
        else:
            parts.append("<p>No job descriptions yet.</p>")
    else:
        parts.append("<p>No job descriptions folder yet.</p>")

    parts.append('<p><a href="/">Home</a></p>')
    return "".join(parts)


@app.get("/download/letters/<path:filename>")
def download_letter(filename):
    return send_from_directory(LETTERS_DIR, filename, as_attachment=True)


@app.get("/download/job_descriptions/<path:filename>")
def download_job_description(filename):
    return send_from_directory(JOB_DESCRIPTIONS_DIR, filename, as_attachment=True)
