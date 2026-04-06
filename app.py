from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import pandas as pd
import os
import re
import zipfile
import shutil
from datetime import datetime, timedelta
from PyPDF2 import PdfMerger

from oig_screenshot import capture_oig, close_oig_session
from cna_screenshot import capture_cna, close_cna_session
from adverse_screenshot import capture_adverse, close_adverse_session

app = Flask(__name__)
CORS(app, resources={r"/api/*": {"origins": "*"}})

UPLOAD_FOLDER = "uploads"
RUNS_FOLDER = "runs"
BACKEND_BASE_URL = "https://workforcecomply-backend-docker.onrender.com"

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


def create_results_excel(employee_results, output_path):
    flagged_rows = [e for e in employee_results if e["flagged"]]

    if flagged_rows:
        results_df = pd.DataFrame(flagged_rows)
    else:
        results_df = pd.DataFrame(columns=["First Name", "Last Name", "SSN", "Issues"])

    summary_lines = []
    for e in employee_results:
        if e["issues"]:
            for issue in e["issues"]:
                summary_lines.append(f"{e['First Name']} {e['Last Name']} - {issue}")

    if summary_lines:
        summary_df = pd.DataFrame({"Issues": summary_lines})
    else:
        summary_df = pd.DataFrame({"Issues": ["No issues found."]})

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        results_df.to_excel(writer, index=False, sheet_name="Flagged Employees")
        summary_df.to_excel(writer, index=False, sheet_name="Summary")


@app.route("/", methods=["GET"])
def home():
    return jsonify({"status": "ok", "message": "WorkforceComply backend is running"})


@app.route("/api/run-checks", methods=["POST"])
def run_checks():
    cleanup_old_runs(RUNS_FOLDER)

    file = request.files.get("file")
    if not file or file.filename == "":
        return jsonify({"error": "No file selected"}), 400

    run_id = datetime.now().strftime("%Y%m%d_%H%M%S")
    run_folder = os.path.join(RUNS_FOLDER, run_id)

    oig_folder = os.path.join(run_folder, "OIG_Report")
    cna_folder = os.path.join(run_folder, "CNA_Report")
    adverse_folder = os.path.join(run_folder, "Adverse_Actions_Report")

    os.makedirs(oig_folder, exist_ok=True)
    os.makedirs(cna_folder, exist_ok=True)
    os.makedirs(adverse_folder, exist_ok=True)

    upload_path = os.path.join(UPLOAD_FOLDER, f"{run_id}_{file.filename}")
    file.save(upload_path)

    try:
        df = pd.read_excel(upload_path)
    except Exception as e:
        return jsonify({"error": f"Could not read Excel file: {e}"}), 400

    required_columns = ["First Name", "Last Name", "SSN"]
    missing = [col for col in required_columns if col not in df.columns]
    if missing:
        return jsonify({"error": f"Missing required column(s): {', '.join(missing)}"}), 400

    employee_results = []
    oig_paths = []
    cna_paths = []
    adverse_paths = []

    try:
        for _, row in df.iterrows():
            first = safe_text(row["First Name"])
            last = safe_text(row["Last Name"])
            ssn_raw = row["SSN"]
            ssn = clean_ssn(ssn_raw)

            employee = {
                "First Name": first,
                "Last Name": last,
                "SSN": ssn,
                "issues": [],
                "flagged": False
            }

            # OIG
            oig_result = capture_oig(first, last, oig_folder)
            oig_pdf = oig_result.get("pdf_path")

            if oig_pdf:
                oig_paths.append(oig_pdf)
            else:
                error = oig_result.get("error")
                if error:
                    employee["issues"].append(f"REVIEW NEEDED - OIG ERROR: {error}")
                else:
                    employee["issues"].append("REVIEW NEEDED - OIG PROOF MISSING")

            # CNA
            if len(ssn) != 9:
                employee["issues"].append("ERROR - INVALID SSN FOR CNA")
            else:
                cna_result = capture_cna(ssn, cna_folder)
                cna_pdf = cna_result.get("pdf_path")

                if cna_pdf:
                    cna_paths.append(cna_pdf)
                else:
                    error = cna_result.get("error")
                    if error:
                        employee["issues"].append(f"REVIEW NEEDED - CNA ERROR: {error}")
                    else:
                        employee["issues"].append("REVIEW NEEDED - CNA PROOF MISSING")

            # ADVERSE
            if len(ssn) != 9:
                employee["issues"].append("ERROR - INVALID SSN FOR ADVERSE")
            else:
                adverse_result = capture_adverse(first, last, ssn, adverse_folder)
                adverse_pdf = adverse_result.get("pdf_path")

                if adverse_pdf:
                    adverse_paths.append(adverse_pdf)
                else:
                    error = adverse_result.get("error")
                    if error:
                        employee["issues"].append(f"REVIEW NEEDED - ADVERSE ERROR: {error}")
                    else:
                        employee["issues"].append("REVIEW NEEDED - ADVERSE PROOF MISSING")

            if employee["issues"]:
                employee["flagged"] = True

            employee_results.append(employee)

    finally:
        close_oig_session()
        close_cna_session()
        close_adverse_session()

    oig_merged = os.path.join(oig_folder, "OIG_Merged.pdf")
    cna_merged = os.path.join(cna_folder, "CNA_Merged.pdf")
    adverse_merged = os.path.join(adverse_folder, "Adverse_Actions_Merged.pdf")

    merge_pdfs(oig_paths, oig_merged)
    merge_pdfs(cna_paths, cna_merged)
    merge_pdfs(adverse_paths, adverse_merged)

    summary_path = os.path.join(run_folder, "SUMMARY.txt")
    with open(summary_path, "w", encoding="utf-8") as f:
        f.write("WORKFORCECOMPLY SUMMARY\n\n")

        issues_exist = False
        for e in employee_results:
            for issue in e["issues"]:
                issues_exist = True
                f.write(f"{e['First Name']} {e['Last Name']} - {issue}\n")

        if not issues_exist:
            f.write("All employees processed successfully.\n")

    results_excel_path = os.path.join(run_folder, "Results.xlsx")
    create_results_excel(employee_results, results_excel_path)

    zip_path = os.path.join(run_folder, "output.zip")
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
        if os.path.exists(oig_merged):
            z.write(oig_merged, "OIG_Report/OIG_Merged.pdf")
        if os.path.exists(cna_merged):
            z.write(cna_merged, "CNA_Report/CNA_Merged.pdf")
        if os.path.exists(adverse_merged):
            z.write(adverse_merged, "Adverse_Actions_Report/Adverse_Actions_Merged.pdf")
        if os.path.exists(results_excel_path):
            z.write(results_excel_path, "Results.xlsx")
        z.write(summary_path, "SUMMARY.txt")

    total = len(employee_results)
    flagged = sum(1 for e in employee_results if e["flagged"])
    clear = total - flagged

    return jsonify({
        "summary": {
            "total_employees": total,
            "clear_count": clear,
            "attention_needed": flagged
        },
        "downloads": {
            "combined_pdf_url": f"{BACKEND_BASE_URL}/api/download/{run_id}/zip",
            "individual_zip_url": f"{BACKEND_BASE_URL}/api/download/{run_id}/zip",
            "results_excel_url": f"{BACKEND_BASE_URL}/api/download/{run_id}/results-excel"
        }
    })


@app.route("/api/download/<run_id>/zip", methods=["GET"])
def download_zip(run_id):
    zip_path = os.path.join(RUNS_FOLDER, run_id, "output.zip")
    if not os.path.exists(zip_path):
        return jsonify({"error": "ZIP file not found"}), 404
    return send_file(zip_path, as_attachment=True)


@app.route("/api/download/<run_id>/summary", methods=["GET"])
def download_summary(run_id):
    summary_path = os.path.join(RUNS_FOLDER, run_id, "SUMMARY.txt")
    if not os.path.exists(summary_path):
        return jsonify({"error": "Summary file not found"}), 404
    return send_file(summary_path, as_attachment=True)


@app.route("/api/download/<run_id>/results-excel", methods=["GET"])
def download_results_excel(run_id):
    results_excel_path = os.path.join(RUNS_FOLDER, run_id, "Results.xlsx")
    if not os.path.exists(results_excel_path):
        return jsonify({"error": "Results Excel file not found"}), 404
    return send_file(results_excel_path, as_attachment=True)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
