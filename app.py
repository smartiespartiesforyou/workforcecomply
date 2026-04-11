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
CORS(app, resources={r"/api/*": {"origins": "*"}})

UPLOAD_FOLDER = "uploads"
RUNS_FOLDER = "runs"
BACKEND_BASE_URL = "https://workforcecomply-backend-docker.onrender.com"

COMBINED_WORKERS = 3

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RUNS_FOLDER, exist_ok=True)


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


def create_results_excel(employee_results, output_path, mode="combined"):
    flagged_rows = []

    for e in employee_results:
        if e["flagged"]:
            flagged_rows.append({
                "First Name": e["First Name"],
                "Last Name": e["Last Name"],
                "SSN": e["SSN"],
                "Status": "Attention Required"
            })

    df = pd.DataFrame(flagged_rows)
    df.to_excel(output_path, index=False)


def build_zip(run_folder, zip_path, include_folders):
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
        results_path = os.path.join(run_folder, "Results.xlsx")
        if os.path.exists(results_path):
            z.write(results_path, "Results.xlsx")

        for folder_name in include_folders:
            folder_path = os.path.join(run_folder, folder_name)
            if not os.path.exists(folder_path):
                continue

            for root, dirs, files in os.walk(folder_path):
                for file_name in files:
                    full_path = os.path.join(root, file_name)
                    relative_path = os.path.relpath(full_path, run_folder)
                    z.write(full_path, relative_path)


def run_oig_safe(first, last, save_folder):
    return capture_oig(first, last, save_folder)


def run_cna_safe(ssn, save_folder):
    return capture_cna(ssn, save_folder)


def run_adverse_safe(first, last, ssn, save_folder):
    return capture_adverse(first, last, ssn, save_folder)


def prepare_upload(file, run_id):
    upload_name = os.path.basename(file.filename)
    upload_path = os.path.join(UPLOAD_FOLDER, f"{run_id}_{upload_name}")
    file.save(upload_path)
    return upload_path


def read_input_dataframe(upload_path):
    df = pd.read_excel(upload_path)

    required_columns = ["First Name", "Last Name", "SSN"]
    for col in required_columns:
        if col not in df.columns:
            raise ValueError(f"Missing required column: {col}")

    return df


def process_combined_run(df, run_folder):
    oig_folder = os.path.join(run_folder, "OIG_Report")
    cna_folder = os.path.join(run_folder, "CNA_Report")
    adverse_folder = os.path.join(run_folder, "Adverse_Actions_Report")

    os.makedirs(oig_folder, exist_ok=True)
    os.makedirs(cna_folder, exist_ok=True)
    os.makedirs(adverse_folder, exist_ok=True)

    employee_results = []
    oig_paths = []
    cna_paths = []
    adverse_paths = []

    def process_single(row):
        first = safe_text(row["First Name"])
        last = safe_text(row["Last Name"])
        ssn = clean_ssn(row["SSN"])

        employee = {
            "First Name": first,
            "Last Name": last,
            "SSN": ssn,
            "issues": [],
            "flagged": False
        }

        try:
            oig = run_oig_safe(first, last, oig_folder)
            if oig.get("pdf_path"):
                oig_paths.append(oig["pdf_path"])
        except:
            employee["issues"].append("OIG ERROR")

        if len(ssn) == 9:
            try:
                cna = run_cna_safe(ssn, cna_folder)
                if cna.get("pdf_path"):
                    cna_paths.append(cna["pdf_path"])
            except:
                employee["issues"].append("CNA ERROR")
        else:
            employee["issues"].append("INVALID SSN")

        if len(ssn) == 9:
            try:
                adv = run_adverse_safe(first, last, ssn, adverse_folder)
                if adv.get("pdf_path"):
                    adverse_paths.append(adv["pdf_path"])
            except:
                employee["issues"].append("ADVERSE ERROR")

        if employee["issues"]:
            employee["flagged"] = True

        return employee

    with ThreadPoolExecutor(max_workers=COMBINED_WORKERS) as executor:
        results = executor.map(process_single, [row for _, row in df.iterrows()])
        for r in results:
            employee_results.append(r)

    close_oig_session()
    close_cna_session()
    close_adverse_session()

    merge_pdfs(oig_paths, os.path.join(oig_folder, "OIG_Merged.pdf"))
    merge_pdfs(cna_paths, os.path.join(cna_folder, "CNA_Merged.pdf"))
    merge_pdfs(adverse_paths, os.path.join(adverse_folder, "Adverse_Actions_Merged.pdf"))

    results_excel_path = os.path.join(run_folder, "Results.xlsx")
    create_results_excel(employee_results, results_excel_path)

    zip_path = os.path.join(run_folder, "output.zip")
    build_zip(run_folder, zip_path, ["OIG_Report", "CNA_Report", "Adverse_Actions_Report"])

    return employee_results


@app.route("/api/run-checks", methods=["POST"])
def run_checks():
    cleanup_old_runs(RUNS_FOLDER)

    file = request.files.get("file")
    if not file:
        return jsonify({"error": "No file"}), 400

    run_id = datetime.now().strftime("%Y%m%d_%H%M%S")
    run_folder = os.path.join(RUNS_FOLDER, run_id)
    os.makedirs(run_folder)

    upload = prepare_upload(file, run_id)
    df = read_input_dataframe(upload)

    results = process_combined_run(df, run_folder)

    total = len(results)
    flagged = sum(1 for e in results if e["flagged"])

    return jsonify({
        "summary": {
            "total_employees": total,
            "clear_count": total - flagged,
            "attention_needed": flagged
        },
        "downloads": {
            "combined_pdf_url": f"{BACKEND_BASE_URL}/api/download/{run_id}/zip",
            "results_excel_url": f"{BACKEND_BASE_URL}/api/download/{run_id}/results-excel"
        }
    })


@app.route("/api/download/<run_id>/zip")
def download_zip(run_id):
    return send_file(os.path.join(RUNS_FOLDER, run_id, "output.zip"), as_attachment=True)


@app.route("/api/download/<run_id>/results-excel")
def download_excel(run_id):
    return send_file(os.path.join(RUNS_FOLDER, run_id, "Results.xlsx"), as_attachment=True)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
