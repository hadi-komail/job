from flask import Flask
import subprocess
import sys
import traceback
from pathlib import Path

app = Flask(__name__)
BASE_DIR = Path(__file__).resolve().parent


@app.get("/")
def home():
    return """
    <h1>Job Automation</h1>
    <p>Service is running.</p>
    <p><a href="/run">Run the job search</a></p>
    """


@app.get("/run")
def run_script():
    try:
        result = subprocess.run(
            [sys.executable, str(BASE_DIR / "main.py")],
            capture_output=True,
            text=True,
            cwd=BASE_DIR,
            timeout=900,
        )

        output = result.stdout + "\n" + result.stderr
        return f"<pre>{output}</pre>"

    except subprocess.TimeoutExpired as exc:
        return (
            "<pre>"
            f"Script timed out after {exc.timeout} seconds.\n\n"
            f"Partial stdout:\n{exc.stdout or ''}\n\n"
            f"Partial stderr:\n{exc.stderr or ''}"
            "</pre>",
            500,
        )
    except Exception:
        return f"<pre>{traceback.format_exc()}</pre>", 500
