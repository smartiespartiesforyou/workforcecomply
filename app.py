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


def snapshot_files(folder):
    return {
        os.path.join(folder, f): os.path.getmtime(os.path.join(folder, f))
        for f in os.listdir(folder)
        if f.lower().endswith(".pdf")
    }


def get_new_file(before, folder):
    after = snapshot_files(folder)
    new_files = [f for f in after if f not in before]
    if not new_files:
        return None
    return sorted(new_files, key=lambda x: os.path.getmtime(x), reverse=True)[0]


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


def create_results_excel(df, summary_lines, output_path):
    results_df = df.copy()

    if summary_lines:
        summary_df = pd.DataFrame({"Issues": summary_lines})
    else:
        summary_df = pd.DataFrame({"Issues": ["All employees processed successfully."]})

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        results_df.to_excel(writer, index=False, sheet_name="Uploaded Employees")
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

    summary_lines = []
    oig_temp_pdf_paths = []
    cna_temp_pdf_paths = []
    adverse_temp_pdf_paths = []

    try:
        for _, row in df.iterrows():
            first = safe_text(row["First Name"])
            last = safe_text(row["Last Name"])
            full_name = f"{first} {last}".strip()

            before_oig = snapshot_files(oig_folder)
            result = capture_oig(first, last, oig_folder)
            new_oig_pdf = get_new_file(before_oig, oig_folder)

            if new_oig_pdf:
                oig_temp_pdf_paths.append(new_oig_pdf)
            else:
                error_text = result.get("error") if isinstance(result, dict) else None
                if error_text:
                    summary_lines.append(f"{full_name} - REVIEW NEEDED - OIG ERROR: {error_text}")
                else:
                    summary_lines.append(f"{full_name} - REVIEW NEEDED - OIG PROOF MISSING")
    finally:
        close_oig_session()

    try:
        for _, row in df.iterrows():
            first = safe_text(row["First Name"])
            last = safe_text(row["Last Name"])
            ssn = clean_ssn(row["SSN"])
            full_name = f"{first} {last}".strip()

            if len(ssn) != 9:
                summary_lines.append(f"{full_name} - ERROR - INVALID SSN FOR CNA")
                continue

            before_cna = snapshot_files(cna_folder)
            result = capture_cna(ssn, cna_folder)
            new_cna_pdf = get_new_file(before_cna, cna_folder)

            if new_cna_pdf:
                cna_temp_pdf_paths.append(new_cna_pdf)
            else:
                error_text = result.get("error") if isinstance(result, dict) else None
                if error_text:
                    summary_lines.append(f"{full_name} - REVIEW NEEDED - CNA ERROR: {error_text}")
                else:
                    summary_lines.append(f"{full_name} - REVIEW NEEDED - CNA PROOF MISSING")
    finally:
        close_cna_session()

    try:
        for _, row in df.iterrows():
            first = safe_text(row["First Name"])
            last = safe_text(row["Last Name"])
            ssn = clean_ssn(row["SSN"])
            full_name = f"{first} {last}".strip()

            if len(ssn) != 9:
                summary_lines.append(f"{full_name} - ERROR - INVALID SSN FOR ADVERSE")
                continue

            before_adverse = snapshot_files(adverse_folder)
            result = capture_adverse(first, last, ssn, adverse_folder)
            new_adverse_pdf = get_new_file(before_adverse, adverse_folder)

            if new_adverse_pdf:
                adverse_temp_pdf_paths.append(new_adverse_pdf)
            else:
                error_text = result.get("error") if isinstance(result, dict) else None
                if error_text:
                    summary_lines.append(f"{full_name} - REVIEW NEEDED - ADVERSE ERROR: {error_text}")
                else:
                    summary_lines.append(f"{full_name} - REVIEW NEEDED - ADVERSE PROOF MISSING")
    finally:
        close_adverse_session()

    oig_merged_path = os.path.join(oig_folder, "OIG_Merged.pdf")
    cna_merged_path = os.path.join(cna_folder, "CNA_Merged.pdf")
    adverse_merged_path = os.path.join(adverse_folder, "Adverse_Actions_Merged.pdf")

    merge_pdfs(oig_temp_pdf_paths, oig_merged_path)
    merge_pdfs(cna_temp_pdf_paths, cna_merged_path)
    merge_pdfs(adverse_temp_pdf_paths, adverse_merged_path)

    summary_path = os.path.join(run_folder, "SUMMARY.txt")
    with open(summary_path, "w", encoding="utf-8") as f:
        f.write("WORKFORCECOMPLY SUMMARY\n\n")

        if summary_lines:
            f.write("REVIEW REQUIRED / ERRORS:\n\n")
            for line in summary_lines:
                f.write(line + "\n")
        else:
            f.write("All employees processed successfully.\n")

    results_excel_path = os.path.join(run_folder, "Results.xlsx")
    create_results_excel(df, summary_lines, results_excel_path)

    zip_path = os.path.join(run_folder, "output.zip")
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
        if os.path.exists(oig_merged_path):
            z.write(oig_merged_path, "OIG_Report/OIG_Merged.pdf")
        if os.path.exists(cna_merged_path):
            z.write(cna_merged_path, "CNA_Report/CNA_Merged.pdf")
        if os.path.exists(adverse_merged_path):
            z.write(adverse_merged_path, "Adverse_Actions_Report/Adverse_Actions_Merged.pdf")
        if os.path.exists(results_excel_path):
            z.write(results_excel_path, "Results.xlsx")
        z.write(summary_path, "SUMMARY.txt")

    total_employees = len(df)
    attention_needed = len(summary_lines)
    clear_count = total_employees - attention_needed

    return jsonify({
        "summary": {
            "total_employees": total_employees,
            "clear_count": clear_count,
            "attention_needed": attention_needed
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
