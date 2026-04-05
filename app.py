from flask import Flask, render_template, request, send_file
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

UPLOAD_FOLDER = "uploads"
RUNS_FOLDER = "runs"

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


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        cleanup_old_runs(RUNS_FOLDER)

        file = request.files.get("file")
        if not file or file.filename == "":
            return "No file selected", 400

        run_id = datetime.now().strftime("%Y%m%d_%H%M%S")
        run_folder = os.path.join(RUNS_FOLDER, run_id)

        oig_folder = os.path.join(run_folder, "OIG_Report")
        cna_folder = os.path.join(run_folder, "CNA_Report")
        adverse_folder = os.path.join(run_folder, "Adverse_Actions_Report")

        os.makedirs(oig_folder, exist_ok=True)
        os.makedirs(cna_folder, exist_ok=True)
        os.makedirs(adverse_folder, exist_ok=True)

        upload_path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(upload_path)

        try:
            df = pd.read_excel(upload_path)
        except Exception as e:
            return f"Could not read Excel file: {e}", 400

        required_columns = ["First Name", "Last Name", "SSN"]
        missing = [col for col in required_columns if col not in df.columns]
        if missing:
            return f"Missing required column(s): {', '.join(missing)}", 400

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
                capture_oig(first, last, oig_folder)
                new_oig_pdf = get_new_file(before_oig, oig_folder)

                if new_oig_pdf:
                    oig_temp_pdf_paths.append(new_oig_pdf)
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
                capture_cna(ssn, cna_folder)
                new_cna_pdf = get_new_file(before_cna, cna_folder)

                if new_cna_pdf:
                    cna_temp_pdf_paths.append(new_cna_pdf)
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

                before_adverse = snapshot_files(adverse_folder)
                capture_adverse(first, last, ssn, adverse_folder)
                new_adverse_pdf = get_new_file(before_adverse, adverse_folder)

                if new_adverse_pdf:
                    adverse_temp_pdf_paths.append(new_adverse_pdf)
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

        for pdf in oig_temp_pdf_paths:
            try:
                os.remove(pdf)
            except Exception:
                pass

        for pdf in cna_temp_pdf_paths:
            try:
                os.remove(pdf)
            except Exception:
                pass

        for pdf in adverse_temp_pdf_paths:
            try:
                os.remove(pdf)
            except Exception:
                pass

        summary_path = os.path.join(run_folder, "SUMMARY.txt")
        with open(summary_path, "w", encoding="utf-8") as f:
            f.write("WORKFORCECOMPLY SUMMARY\n\n")

            if summary_lines:
                f.write("REVIEW REQUIRED / ERRORS:\n\n")
                for line in summary_lines:
                    f.write(line + "\n")
            else:
                f.write("All employees processed successfully.\n")

        zip_path = os.path.join(run_folder, "output.zip")
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
            if os.path.exists(oig_merged_path):
                z.write(oig_merged_path, "OIG_Report/OIG_Merged.pdf")
            if os.path.exists(cna_merged_path):
                z.write(cna_merged_path, "CNA_Report/CNA_Merged.pdf")
            if os.path.exists(adverse_merged_path):
                z.write(adverse_merged_path, "Adverse_Actions_Report/Adverse_Actions_Merged.pdf")
            z.write(summary_path, "SUMMARY.txt")

        return send_file(zip_path, as_attachment=True)

    return render_template("index.html")


if __name__ == "__main__":
    app.run(debug=True)