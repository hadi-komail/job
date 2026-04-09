from flask import Flask
import subprocess
import sys
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
    result = subprocess.run(
        [sys.executable, str(BASE_DIR / "main.py")],
        capture_output=True,
        text=True,
        cwd=BASE_DIR,
    )

    output = result.stdout + "\n" + result.stderr
    return f"<pre>{output}</pre>"
