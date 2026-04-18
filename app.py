from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import pandas as pd
import os
import re
import zipfile
import shutil
from datetime import datetime, timedelta
from concurrent.futures import ThreadPoolExecutor
from PyPDF2 import PdfMerger

from oig_screenshot import capture_oig, close_oig_session
from cna_screenshot import capture_cna, close_cna_session
from adverse_screenshot import capture_adverse, close_adverse_session

app = Flask(__name__)

# 🔐 LOCK CORS TO YOUR SITE ONLY
CORS(app, resources={r"/api/*": {"origins": ["https://www.workforcecomply.com"]}})

# 🔐 API KEY (SET IN RENDER ENV VARIABLES)
API_KEY = os.environ.get("API_KEY")

UPLOAD_FOLDER = "uploads"
RUNS_FOLDER = "runs"
BACKEND_BASE_URL = "https://workforcecomply-backend-docker.onrender.com"

COMBINED_WORKERS = 3

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RUNS_FOLDER, exist_ok=True)


# 🔐 AUTH CHECK FUNCTION
def require_api_key():
    key = request.headers.get("x-api-key")
    if not key or key != API_KEY:
        return False
    return True


def safe_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def clean_ssn(value):
    return re.sub(r"\D", "", safe_text(value))


def cleanup_old_runs(folder, days=2):
    now = datetime.now()
    for name in os.listdir(folder):
        path = os.path.join(folder, name)
        if os.path.isdir(path):
            try:
                created = datetime.fromtimestamp(os.path.getctime(path))
                if now - created > timedelta(days=days):
                    shutil.rmtree(path)
            except Exception:
                pass


def merge_pdfs(pdf_paths, output_path):
    valid_paths = [p for p in pdf_paths if p and os.path.exists(p)]
    if not valid_paths:
        return None

    merger = PdfMerger()
    for pdf_path in valid_paths:
        merger.append(pdf_path)
    merger.write(output_path)
    merger.close()
    return output_path


# 🔐 PROTECT ROUTES
def protect():
    if not require_api_key():
        return jsonify({"error": "Unauthorized"}), 401


@app.route("/api/run-checks", methods=["POST"])
def run_checks():
    auth = protect()
    if auth:
        return auth

    cleanup_old_runs(RUNS_FOLDER)

    file = request.files.get("file")
    if not file or file.filename == "":
        return jsonify({"error": "No file selected"}), 400

    run_id = datetime.now().strftime("%Y%m%d_%H%M%S")
    run_folder = os.path.join(RUNS_FOLDER, run_id)
    os.makedirs(run_folder, exist_ok=True)

    upload_path = os.path.join(UPLOAD_FOLDER, f"{run_id}_{file.filename}")
    file.save(upload_path)

    try:
        df = pd.read_excel(upload_path)
    except Exception:
        return jsonify({"error": "Invalid Excel file"}), 400

    # KEEP YOUR EXISTING LOGIC
    employee_results = []  # (rest unchanged for now)

    return jsonify({"status": "secured"})


@app.route("/api/download/<run_id>/zip", methods=["GET"])
def download_zip(run_id):
    auth = protect()
    if auth:
        return auth

    run_folder = os.path.join(RUNS_FOLDER, run_id)

    if not os.path.exists(run_folder):
        return jsonify({"error": "Not found"}), 404

    zip_file = None
    for f in os.listdir(run_folder):
        if f.endswith(".zip"):
            zip_file = os.path.join(run_folder, f)
            break

    if not zip_file:
        return jsonify({"error": "Not found"}), 404

    return send_file(zip_file, as_attachment=True)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
