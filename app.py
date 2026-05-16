from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from werkzeug.utils import secure_filename
import pandas as pd
import os
import re
import zipfile
import shutil
import secrets
import urllib.request
from datetime import datetime, timedelta
from concurrent.futures import ThreadPoolExecutor
from PyPDF2 import PdfMerger
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors

from oig_screenshot import capture_oig, close_oig_session
from cna_screenshot import capture_cna, close_cna_session
from adverse_screenshot import capture_adverse, close_adverse_session

app = Flask(__name__)

ALLOWED_ORIGINS = [
    "https://www.workforcecomply.com",
    "https://workforcecomply.com"
]

CORS(
    app,
    resources={r"/api/*": {"origins": ALLOWED_ORIGINS}},
)

UPLOAD_FOLDER = "uploads"
RUNS_FOLDER = "runs"
BACKEND_BASE_URL = "https://workforcecomply-backend-docker.onrender.com"
API_KEY = os.environ.get("API_KEY", "").strip()

COMBINED_WORKERS = 3
ALLOWED_EXTENSIONS = {".xlsx", ".xls"}
MAX_UPLOAD_MB = 10
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_MB * 1024 * 1024

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RUNS_FOLDER, exist_ok=True)


def safe_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def clean_ssn(value):
    return re.sub(r"\D", "", safe_text(value))


def allowed_file(filename):
    ext = os.path.splitext(filename or "")[1].lower()
    return ext in ALLOWED_EXTENSIONS


def generate_run_id():
    return secrets.token_urlsafe(24)


def cleanup_old_runs(folder, hours=4):
    now = datetime.now()
    for name in os.listdir(folder):
        path = os.path.join(folder, name)
        if os.path.isdir(path):
            try:
                created = datetime.fromtimestamp(os.path.getctime(path))
                if now - created > timedelta(hours=hours):
                    shutil.rmtree(path)
            except Exception:
                pass


def cleanup_old_uploads(folder, hours=4):
    now = datetime.now()
    for name in os.listdir(folder):
        path = os.path.join(folder, name)
        if os.path.isfile(path):
            try:
                created = datetime.fromtimestamp(os.path.getctime(path))
                if now - created > timedelta(hours=hours):
                    os.remove(path)
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


def normalize_status(issue_list, check_name):
    relevant = [issue for issue in issue_list if check_name in issue.upper()]

    if not relevant:
        return "Clear"

    text = " | ".join(relevant).upper()

    if "INVALID SSN" in text:
        return "Invalid SSN"
    if "MATCH" in text:
        return "Match Found"
    if "FOUND" in text:
        return "Match Found"
    if "NOT ACTIVE" in text:
        return "Not Active"
    if "PROOF MISSING" in text:
        return "Proof Missing"
    if "ERROR" in text:
        return "Error"
    if "REVIEW" in text:
        return "Review Needed"

    return "Review Needed"


def create_results_excel(employee_results, output_path, mode="combined"):
    flagged_rows = []

    for e in employee_results:
        if e["flagged"]:
            if mode == "oig":
                flagged_rows.append({
                    "First Name": e["First Name"],
                    "Last Name": e["Last Name"],
                    "SSN": e["SSN"],
                    "OIG": normalize_status(e["issues"], "OIG"),
                    "Status": "Attention Required"
                })
            elif mode == "cna":
                flagged_rows.append({
                    "First Name": e["First Name"],
                    "Last Name": e["Last Name"],
                    "SSN": e["SSN"],
                    "CNA": normalize_status(e["issues"], "CNA"),
                    "Status": "Attention Required"
                })
            elif mode == "adverse":
                flagged_rows.append({
                    "First Name": e["First Name"],
                    "Last Name": e["Last Name"],
                    "SSN": e["SSN"],
                    "DSW Result": normalize_status(e["issues"], "ADVERSE"),
                    "Issue Details": " | ".join(e["issues"]),
                    "Status": "Attention Required"
                })
            else:
                flagged_rows.append({
                    "First Name": e["First Name"],
                    "Last Name": e["Last Name"],
                    "SSN": e["SSN"],
                    "OIG": normalize_status(e["issues"], "OIG"),
                    "CNA": normalize_status(e["issues"], "CNA"),
                    "Adverse": normalize_status(e["issues"], "ADVERSE"),
                    "Status": "Attention Required"
                })

    if flagged_rows:
        df = pd.DataFrame(flagged_rows)
    else:
        if mode == "oig":
            df = pd.DataFrame(columns=[
                "First Name", "Last Name", "SSN", "OIG", "Status"
            ])
        elif mode == "cna":
            df = pd.DataFrame(columns=[
                "First Name", "Last Name", "SSN", "CNA", "Status"
            ])
        elif mode == "adverse":
            df = pd.DataFrame(columns=[
                "First Name", "Last Name", "SSN", "DSW Result", "Issue Details", "Status"
            ])
        else:
            df = pd.DataFrame(columns=[
                "First Name", "Last Name", "SSN", "OIG", "CNA", "Adverse", "Status"
            ])

    df.to_excel(output_path, index=False)


def build_zip(run_folder, zip_path, include_folders, excel_filename):
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
        results_path = os.path.join(run_folder, excel_filename)
        if os.path.exists(results_path):
            z.write(results_path, excel_filename)

        for folder_name in include_folders:
            folder_path = os.path.join(run_folder, folder_name)
            if not os.path.exists(folder_path):
                continue

            for root, dirs, files in os.walk(folder_path):
                for file_name in files:
                    full_path = os.path.join(root, file_name)
                    relative_path = os.path.relpath(full_path, run_folder)
                    z.write(full_path, relative_path)


def run_oig_safe(first, last, ssn, save_folder):
    return capture_oig(first, last, ssn, save_folder)


def run_cna_safe(first, last, ssn, save_folder):
    return capture_cna(first, last, ssn, save_folder)


def run_adverse_safe(first, last, ssn, save_folder):
    return capture_adverse(first, last, ssn, save_folder)


def require_api_key():
    expected = API_KEY
    provided = request.headers.get("x-api-key", "").strip()

    if not expected:
        return jsonify({"error": "Server configuration error"}), 500

    if not provided or provided != expected:
        return jsonify({"error": "Unauthorized"}), 401

    return None


def prepare_upload(file, run_folder):
    original_name = secure_filename(file.filename or "")
    ext = os.path.splitext(original_name)[1].lower()

    if ext not in ALLOWED_EXTENSIONS:
        raise ValueError("Only Excel files (.xlsx or .xls) are allowed")

    upload_path = os.path.join(run_folder, f"input{ext}")
    file.save(upload_path)
    return upload_path


def read_input_dataframe(upload_path):
    df = pd.read_excel(upload_path)

    required_columns = ["First Name", "Last Name", "SSN"]
    missing = [col for col in required_columns if col not in df.columns]
    if missing:
        raise ValueError(f"Missing required column(s): {', '.join(missing)}")

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

    try:
        def process_employee(row):
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

            try:
                oig_result = run_oig_safe(first, last, ssn, oig_folder)
                oig_pdf = oig_result.get("pdf_path")
                oig_status = safe_text(oig_result.get("oig_status", "")).lower()
                oig_match_found = bool(oig_result.get("oig_match_found", False))

                if oig_pdf:
                    oig_paths.append(oig_pdf)
                else:
                    error = oig_result.get("error")
                    if error:
                        employee["issues"].append("REVIEW NEEDED - OIG ERROR")
                    else:
                        employee["issues"].append("REVIEW NEEDED - OIG PROOF MISSING")

                if oig_match_found or oig_status in ("match", "found", "review_needed", "name_match"):
                    employee["issues"].append("REVIEW NEEDED - OIG NAME MATCH")

            except Exception:
                employee["issues"].append("REVIEW NEEDED - OIG ERROR")

            if len(ssn) != 9:
                employee["issues"].append("ERROR - INVALID SSN FOR CNA")
            else:
                try:
                    cna_result = run_cna_safe(first, last, ssn, cna_folder)
                    cna_pdf = cna_result.get("pdf_path")
                    cna_status = cna_result.get("cna_result", "")

                    if cna_pdf:
                        cna_paths.append(cna_pdf)

                    if cna_status == "not_active":
                        employee["issues"].append("CNA NOT ACTIVE")
                    elif cna_status == "review_needed":
                        employee["issues"].append("REVIEW NEEDED - CNA REVIEW")
                    elif cna_status == "error":
                        employee["issues"].append("REVIEW NEEDED - CNA ERROR")

                except Exception:
                    employee["issues"].append("REVIEW NEEDED - CNA ERROR")

            if len(ssn) != 9:
                employee["issues"].append("ERROR - INVALID SSN FOR ADVERSE")
            else:
                try:
                    adverse_result = run_adverse_safe(
                        first,
                        last,
                        ssn,
                        adverse_folder
                    )
                    adverse_pdf = adverse_result.get("pdf_path")

                    if adverse_pdf:
                        adverse_paths.append(adverse_pdf)

                    adverse_status = safe_text(
                        adverse_result.get("adverse_result", adverse_result.get("status", ""))
                    ).lower()

                    if adverse_status in ("match", "found", "review_needed"):
                        detail = safe_text(adverse_result.get("detail")) or "ADVERSE ACTION FOUND"
                        employee["issues"].append(f"REVIEW NEEDED - ADVERSE: {detail}")
                    elif adverse_status == "error":
                        employee["issues"].append("REVIEW NEEDED - ADVERSE ERROR")
                    elif not adverse_pdf:
                        employee["issues"].append("REVIEW NEEDED - ADVERSE PROOF MISSING")

                except Exception:
                    employee["issues"].append("REVIEW NEEDED - ADVERSE ERROR")

            if employee["issues"]:
                employee["flagged"] = True

            return employee

        with ThreadPoolExecutor(max_workers=COMBINED_WORKERS) as executor:
            employee_results.extend(
                list(executor.map(process_employee, [row for _, row in df.iterrows()]))
            )

    finally:
        close_oig_session()
        close_cna_session()
        close_adverse_session()

    merge_pdfs(oig_paths, os.path.join(oig_folder, "OIG_Merged.pdf"))
    merge_pdfs(cna_paths, os.path.join(cna_folder, "CNA_Merged.pdf"))
    merge_pdfs(adverse_paths, os.path.join(adverse_folder, "Adverse_Actions_Merged.pdf"))

    results_excel_path = os.path.join(run_folder, "Results.xlsx")
    create_results_excel(employee_results, results_excel_path, mode="combined")

    zip_path = os.path.join(run_folder, "output.zip")
    build_zip(
        run_folder,
        zip_path,
        include_folders=["OIG_Report", "CNA_Report", "Adverse_Actions_Report"],
        excel_filename="Results.xlsx"
    )

    return employee_results


def process_oig_only_run(df, run_folder):
    oig_folder = os.path.join(run_folder, "OIG_Report")
    os.makedirs(oig_folder, exist_ok=True)

    employee_results = []
    oig_paths = []

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

            try:
                oig_result = run_oig_safe(first, last, ssn, oig_folder)
                oig_pdf = oig_result.get("pdf_path")
                oig_status = safe_text(oig_result.get("oig_status", "")).lower()
                oig_match_found = bool(oig_result.get("oig_match_found", False))

                if oig_pdf:
                    oig_paths.append(oig_pdf)
                else:
                    error = oig_result.get("error")
                    if error:
                        employee["issues"].append("REVIEW NEEDED - OIG ERROR")
                    else:
                        employee["issues"].append("REVIEW NEEDED - OIG PROOF MISSING")

                if oig_match_found or oig_status in ("match", "found", "review_needed", "name_match"):
                    employee["issues"].append("REVIEW NEEDED - OIG NAME MATCH")

            except Exception:
                employee["issues"].append("REVIEW NEEDED - OIG ERROR")

            if employee["issues"]:
                employee["flagged"] = True

            employee_results.append(employee)

    finally:
        close_oig_session()

    merge_pdfs(oig_paths, os.path.join(oig_folder, "OIG_Merged.pdf"))

    results_excel_path = os.path.join(run_folder, "OIG_Results.xlsx")
    create_results_excel(employee_results, results_excel_path, mode="oig")

    zip_path = os.path.join(run_folder, "OIG_Report.zip")
    build_zip(
        run_folder,
        zip_path,
        include_folders=["OIG_Report"],
        excel_filename="OIG_Results.xlsx"
    )

    return employee_results


def process_cna_only_run(df, run_folder):
    cna_folder = os.path.join(run_folder, "CNA_Report")
    os.makedirs(cna_folder, exist_ok=True)

    employee_results = []
    cna_paths = []

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

            if len(ssn) != 9:
                employee["issues"].append("ERROR - INVALID SSN FOR CNA")
            else:
                try:
                    cna_result = run_cna_safe(first, last, ssn, cna_folder)
                    cna_pdf = cna_result.get("pdf_path")
                    cna_status = cna_result.get("cna_result", "")

                    if cna_pdf:
                        cna_paths.append(cna_pdf)

                    if cna_status == "not_active":
                        employee["issues"].append("CNA NOT ACTIVE")
                    elif cna_status == "review_needed":
                        employee["issues"].append("REVIEW NEEDED - CNA REVIEW")
                    elif cna_status == "error":
                        employee["issues"].append("REVIEW NEEDED - CNA ERROR")

                except Exception:
                    employee["issues"].append("REVIEW NEEDED - CNA ERROR")

            if employee["issues"]:
                employee["flagged"] = True

            employee_results.append(employee)

    finally:
        close_cna_session()

    merge_pdfs(cna_paths, os.path.join(cna_folder, "CNA_Merged.pdf"))

    results_excel_path = os.path.join(run_folder, "CNA_Results.xlsx")
    create_results_excel(employee_results, results_excel_path, mode="cna")

    zip_path = os.path.join(run_folder, "CNA_Report.zip")
    build_zip(
        run_folder,
        zip_path,
        include_folders=["CNA_Report"],
        excel_filename="CNA_Results.xlsx"
    )

    return employee_results




DSW_CSV_URL = "https://adverseactions.ldh.la.gov/SelSearch/SelSearch/GetCsv"


def normalize_match_text(value):
    text = safe_text(value).upper()
    text = re.sub(r"[^A-Z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def download_dsw_csv(run_folder):
    csv_path = os.path.join(run_folder, "DSW_Source_Adverse_Actions_List.csv")

    try:
        request = urllib.request.Request(
            DSW_CSV_URL,
            headers={
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
                "Accept": "text/csv,application/csv,text/plain,*/*",
                "Referer": "https://adverseactions.ldh.la.gov/SelSearch/SelSearch/Export"
            }
        )

        with urllib.request.urlopen(request, timeout=60) as response:
            csv_bytes = response.read()

        with open(csv_path, "wb") as f:
            f.write(csv_bytes)

        df = pd.read_csv(csv_path, dtype=str, keep_default_na=False)
        return df, csv_path, None

    except Exception as e:
        return None, None, str(e)


def dsw_csv_possible_match(first, last, ssn, dsw_df):
    first_norm = normalize_match_text(first)
    last_norm = normalize_match_text(last)
    ssn_clean = clean_ssn(ssn)

    if dsw_df is None or dsw_df.empty:
        return False, ""

    for _, csv_row in dsw_df.iterrows():
        row_parts = [safe_text(v) for v in csv_row.values]
        row_text_raw = " ".join(row_parts)
        row_text_norm = normalize_match_text(row_text_raw)
        row_ssn = clean_ssn(row_text_raw)

        if ssn_clean and len(ssn_clean) == 9 and ssn_clean in row_ssn:
            return True, row_text_raw[:500]

        if first_norm and last_norm:
            if first_norm in row_text_norm and last_norm in row_text_norm:
                return True, row_text_raw[:500]

    return False, ""




def process_adverse_only_run(df, run_folder):
    adverse_folder = os.path.join(run_folder, "Adverse_Actions_Report")
    os.makedirs(adverse_folder, exist_ok=True)

    employee_results = []
    adverse_paths = []

    dsw_df, csv_path, csv_error = download_dsw_csv(run_folder)

    if csv_error or dsw_df is None:
        print(f"DSW CSV download failed. Continuing with live LDH proof checks only. Error: {csv_error}")
    else:
        print(f"DSW CSV downloaded successfully with {len(dsw_df)} rows. Live LDH proof PDFs will still be captured for each employee.")

    # This function now receives an already-small batch from the frontend workflow.
    # Do NOT split again inside this already-small DSW batch.
    rows = [row for _, row in df.iterrows()]
    print(f"Starting DSW official proof batch with {len(rows)} employee(s)")

    def process_one_row(row):
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

        if len(ssn) != 9:
            employee["issues"].append("ERROR - INVALID SSN FOR ADVERSE")
        else:
            csv_possible_match = False
            csv_detail = ""

            if dsw_df is not None and not dsw_df.empty:
                try:
                    csv_possible_match, csv_detail = dsw_csv_possible_match(first, last, ssn, dsw_df)
                except Exception:
                    csv_possible_match = False
                    csv_detail = ""

            try:
                adverse_result = run_adverse_safe(first, last, ssn, adverse_folder)
                adverse_pdf = adverse_result.get("pdf_path")

                if adverse_pdf:
                    adverse_paths.append(adverse_pdf)

                adverse_status = safe_text(
                    adverse_result.get("adverse_result", adverse_result.get("status", ""))
                ).lower()

                if adverse_status in ("match", "found", "review_needed"):
                    detail = safe_text(adverse_result.get("detail")) or csv_detail or "ADVERSE ACTION FOUND"
                    employee["issues"].append(f"REVIEW NEEDED - ADVERSE: {detail}")
                elif adverse_status == "error":
                    employee["issues"].append("REVIEW NEEDED - ADVERSE ERROR")
                elif not adverse_pdf:
                    employee["issues"].append("REVIEW NEEDED - ADVERSE PROOF MISSING")
                elif csv_possible_match:
                    employee["issues"].append(
                        f"REVIEW NEEDED - CSV POSSIBLE MATCH NEEDS MANUAL REVIEW: {csv_detail}"
                    )

            except Exception:
                employee["issues"].append("REVIEW NEEDED - ADVERSE ERROR")

        if employee["issues"]:
            employee["flagged"] = True

        return employee

    try:
        with ThreadPoolExecutor(max_workers=1) as executor:
            employee_results = list(executor.map(process_one_row, rows))

    finally:
        close_adverse_session()

    merge_pdfs(adverse_paths, os.path.join(adverse_folder, "Adverse_Actions_Merged.pdf"))

    results_excel_path = os.path.join(run_folder, "DSW_Results.xlsx")
    create_results_excel(employee_results, results_excel_path, mode="adverse")

    zip_path = os.path.join(run_folder, "DSW_Report.zip")
    build_zip(
        run_folder,
        zip_path,
        include_folders=["Adverse_Actions_Report"],
        excel_filename="DSW_Results.xlsx"
    )

    return employee_results


def process_single_person_run(first, last, ssn, run_folder):
    oig_folder = os.path.join(run_folder, "OIG_Report")
    cna_folder = os.path.join(run_folder, "CNA_Report")
    adverse_folder = os.path.join(run_folder, "Adverse_Actions_Report")

    os.makedirs(oig_folder, exist_ok=True)
    os.makedirs(cna_folder, exist_ok=True)
    os.makedirs(adverse_folder, exist_ok=True)

    employee = {
        "First Name": first,
        "Last Name": last,
        "SSN": ssn,
        "issues": [],
        "flagged": False
    }

    oig_paths = []
    cna_paths = []
    adverse_paths = []

    try:
        try:
            oig_result = run_oig_safe(first, last, ssn, oig_folder)
            oig_pdf = oig_result.get("pdf_path")
            oig_status = safe_text(oig_result.get("oig_status", "")).lower()
            oig_match_found = bool(oig_result.get("oig_match_found", False))

            if oig_pdf:
                oig_paths.append(oig_pdf)
            else:
                error = oig_result.get("error")
                if error:
                    employee["issues"].append("REVIEW NEEDED - OIG ERROR")
                else:
                    employee["issues"].append("REVIEW NEEDED - OIG PROOF MISSING")

            if oig_match_found or oig_status in ("match", "found", "review_needed", "name_match"):
                employee["issues"].append("REVIEW NEEDED - OIG NAME MATCH")

        except Exception:
            employee["issues"].append("REVIEW NEEDED - OIG ERROR")

        if len(ssn) != 9:
            employee["issues"].append("ERROR - INVALID SSN FOR CNA")
            employee["issues"].append("ERROR - INVALID SSN FOR ADVERSE")
        else:
            try:
                cna_result = run_cna_safe(first, last, ssn, cna_folder)
                cna_pdf = cna_result.get("pdf_path")
                cna_status = cna_result.get("cna_result", "")

                if cna_pdf:
                    cna_paths.append(cna_pdf)

                if cna_status == "not_active":
                    employee["issues"].append("CNA NOT ACTIVE")
                elif cna_status == "review_needed":
                    employee["issues"].append("REVIEW NEEDED - CNA REVIEW")
                elif cna_status == "error":
                    employee["issues"].append("REVIEW NEEDED - CNA ERROR")
                elif not cna_pdf:
                    employee["issues"].append("REVIEW NEEDED - CNA PROOF MISSING")

            except Exception:
                employee["issues"].append("REVIEW NEEDED - CNA ERROR")

            try:
                adverse_result = run_adverse_safe(first, last, ssn, adverse_folder)
                adverse_pdf = adverse_result.get("pdf_path")

                if adverse_pdf:
                    adverse_paths.append(adverse_pdf)

                adverse_status = safe_text(
                    adverse_result.get("adverse_result", adverse_result.get("status", ""))
                ).lower()

                if adverse_status in ("match", "found", "review_needed"):
                    detail = safe_text(adverse_result.get("detail")) or "ADVERSE ACTION FOUND"
                    employee["issues"].append(f"REVIEW NEEDED - ADVERSE: {detail}")
                elif adverse_status == "error":
                    employee["issues"].append("REVIEW NEEDED - ADVERSE ERROR")
                elif not adverse_pdf:
                    employee["issues"].append("REVIEW NEEDED - ADVERSE PROOF MISSING")

            except Exception:
                employee["issues"].append("REVIEW NEEDED - ADVERSE ERROR")

        if employee["issues"]:
            employee["flagged"] = True

    finally:
        close_oig_session()
        close_cna_session()
        close_adverse_session()

    employee_results = [employee]

    oig_merged = merge_pdfs(oig_paths, os.path.join(oig_folder, "OIG_Merged.pdf"))
    cna_merged = merge_pdfs(cna_paths, os.path.join(cna_folder, "CNA_Merged.pdf"))
    adverse_merged = merge_pdfs(adverse_paths, os.path.join(adverse_folder, "Adverse_Actions_Merged.pdf"))

    pdf_packet_paths = [p for p in [oig_merged, cna_merged, adverse_merged] if p]

    safe_first = re.sub(r"[^A-Za-z0-9_-]+", "_", first).strip("_") or "First"
    safe_last = re.sub(r"[^A-Za-z0-9_-]+", "_", last).strip("_") or "Last"

    final_pdf_path = os.path.join(
        run_folder,
        f"Single_Person_Check_{safe_first}_{safe_last}.pdf"
    )

    merge_pdfs(pdf_packet_paths, final_pdf_path)

    results_excel_path = os.path.join(run_folder, "Single_Person_Results.xlsx")
    create_results_excel(employee_results, results_excel_path, mode="combined")

    return employee_results


def make_response(run_id, employee_results, single_pdf=False):
    total = len(employee_results)
    flagged = sum(1 for e in employee_results if e["flagged"])
    clear = total - flagged

    downloads = {
        "zip_url": f"{BACKEND_BASE_URL}/api/download/{run_id}/zip"
    }

    if single_pdf:
        downloads["pdf_url"] = f"{BACKEND_BASE_URL}/api/download/{run_id}/single-pdf"

    return jsonify({
        "summary": {
            "total_employees": total,
            "clear_count": clear,
            "attention_needed": flagged
        },
        "downloads": downloads
    })




def get_dsw_batch_size():
    return 5


def split_dataframe_for_dsw_batches(df):
    df = df.copy()

    df["_LAST_NAME_SORT"] = (
        df["Last Name"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
    )

    df["_FIRST_NAME_SORT"] = (
        df["First Name"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
    )

    df = df.sort_values(by=["_LAST_NAME_SORT", "_FIRST_NAME_SORT"])

    clean_df = df.drop(
        columns=["_LAST_NAME_SORT", "_FIRST_NAME_SORT"],
        errors="ignore"
    ).reset_index(drop=True)

    batch_size = get_dsw_batch_size()
    batches = []

    for start_index in range(0, len(clean_df), batch_size):
        batch_df = clean_df.iloc[start_index:start_index + batch_size].copy()
        batch_number = len(batches) + 1
        batch_id = f"batch-{batch_number:03d}"
        label = f"Batch {batch_number}"

        batches.append({
            "id": batch_id,
            "label": label,
            "count": int(len(batch_df)),
            "df": batch_df
        })

    return batches



def make_batch_response(parent_run_id, batch_id, employee_results):
    total = len(employee_results)
    flagged = sum(1 for e in employee_results if e["flagged"])
    clear = total - flagged

    return jsonify({
        "batch": {
            "id": batch_id,
            "total_employees": total,
            "clear_count": clear,
            "attention_needed": flagged
        },
        "downloads": {
            "zip_url": f"{BACKEND_BASE_URL}/api/download/{parent_run_id}/{batch_id}/zip"
        }
    })


@app.errorhandler(413)
def too_large(_error):
    return jsonify({"error": f"File too large. Maximum size is {MAX_UPLOAD_MB} MB."}), 413


@app.route("/", methods=["GET"])
def home():
    return jsonify({"status": "ok", "message": "WorkforceComply backend is running"})


@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok"})


@app.route("/api/run-checks", methods=["POST"])
def run_checks():
    auth_error = require_api_key()
    if auth_error:
        return auth_error

    cleanup_old_runs(RUNS_FOLDER)
    cleanup_old_uploads(UPLOAD_FOLDER)

    file = request.files.get("file")
    if not file or file.filename == "":
        return jsonify({"error": "No file selected"}), 400

    if not allowed_file(file.filename):
        return jsonify({"error": "Only Excel files (.xlsx or .xls) are allowed"}), 400

    run_id = generate_run_id()
    run_folder = os.path.join(RUNS_FOLDER, run_id)
    os.makedirs(run_folder, exist_ok=True)

    upload_path = None

    try:
        upload_path = prepare_upload(file, run_folder)
        df = read_input_dataframe(upload_path)
        employee_results = process_combined_run(df, run_folder)
        return make_response(run_id, employee_results)
    except ValueError as e:
        if os.path.exists(run_folder):
            shutil.rmtree(run_folder, ignore_errors=True)
        return jsonify({"error": str(e)}), 400
    except Exception:
        if os.path.exists(run_folder):
            shutil.rmtree(run_folder, ignore_errors=True)
        return jsonify({"error": "Processing failed"}), 500
    finally:
        if upload_path and os.path.exists(upload_path):
            try:
                os.remove(upload_path)
            except Exception:
                pass


@app.route("/api/run-single", methods=["POST"])
def run_single_person():
    auth_error = require_api_key()
    if auth_error:
        return auth_error

    cleanup_old_runs(RUNS_FOLDER)
    cleanup_old_uploads(UPLOAD_FOLDER)

    data = request.get_json(silent=True) or request.form

    first = safe_text(data.get("first_name") or data.get("First Name"))
    last = safe_text(data.get("last_name") or data.get("Last Name"))
    ssn = clean_ssn(data.get("ssn") or data.get("SSN"))

    if not first:
        return jsonify({"error": "First name is required"}), 400

    if not last:
        return jsonify({"error": "Last name is required"}), 400

    if not ssn:
        return jsonify({"error": "SSN is required"}), 400

    run_id = generate_run_id()
    run_folder = os.path.join(RUNS_FOLDER, run_id)
    os.makedirs(run_folder, exist_ok=True)

    try:
        employee_results = process_single_person_run(first, last, ssn, run_folder)
        return make_response(run_id, employee_results, single_pdf=True)
    except Exception:
        if os.path.exists(run_folder):
            shutil.rmtree(run_folder, ignore_errors=True)
        return jsonify({"error": "Processing failed"}), 500


@app.route("/api/run-oig", methods=["POST"])
def run_oig_only():
    auth_error = require_api_key()
    if auth_error:
        return auth_error

    cleanup_old_runs(RUNS_FOLDER)
    cleanup_old_uploads(UPLOAD_FOLDER)

    file = request.files.get("file")
    if not file or file.filename == "":
        return jsonify({"error": "No file selected"}), 400

    if not allowed_file(file.filename):
        return jsonify({"error": "Only Excel files (.xlsx or .xls) are allowed"}), 400

    run_id = generate_run_id()
    run_folder = os.path.join(RUNS_FOLDER, run_id)
    os.makedirs(run_folder, exist_ok=True)

    upload_path = None

    try:
        upload_path = prepare_upload(file, run_folder)
        df = read_input_dataframe(upload_path)
        employee_results = process_oig_only_run(df, run_folder)
        return make_response(run_id, employee_results)
    except ValueError as e:
        if os.path.exists(run_folder):
            shutil.rmtree(run_folder, ignore_errors=True)
        return jsonify({"error": str(e)}), 400
    except Exception:
        if os.path.exists(run_folder):
            shutil.rmtree(run_folder, ignore_errors=True)
        return jsonify({"error": "Processing failed"}), 500
    finally:
        if upload_path and os.path.exists(upload_path):
            try:
                os.remove(upload_path)
            except Exception:
                pass


@app.route("/api/run-cna", methods=["POST"])
def run_cna_only():
    auth_error = require_api_key()
    if auth_error:
        return auth_error

    cleanup_old_runs(RUNS_FOLDER)
    cleanup_old_uploads(UPLOAD_FOLDER)

    file = request.files.get("file")
    if not file or file.filename == "":
        return jsonify({"error": "No file selected"}), 400

    if not allowed_file(file.filename):
        return jsonify({"error": "Only Excel files (.xlsx or .xls) are allowed"}), 400

    run_id = generate_run_id()
    run_folder = os.path.join(RUNS_FOLDER, run_id)
    os.makedirs(run_folder, exist_ok=True)

    upload_path = None

    try:
        upload_path = prepare_upload(file, run_folder)
        df = read_input_dataframe(upload_path)
        employee_results = process_cna_only_run(df, run_folder)
        return make_response(run_id, employee_results)
    except ValueError as e:
        if os.path.exists(run_folder):
            shutil.rmtree(run_folder, ignore_errors=True)
        return jsonify({"error": str(e)}), 400
    except Exception:
        if os.path.exists(run_folder):
            shutil.rmtree(run_folder, ignore_errors=True)
        return jsonify({"error": "Processing failed"}), 500
    finally:
        if upload_path and os.path.exists(upload_path):
            try:
                os.remove(upload_path)
            except Exception:
                pass




@app.route("/api/run-adverse-batched-init", methods=["POST"])
def run_adverse_batched_init():
    auth_error = require_api_key()
    if auth_error:
        return auth_error

    cleanup_old_runs(RUNS_FOLDER)
    cleanup_old_uploads(UPLOAD_FOLDER)

    file = request.files.get("file")
    if not file or file.filename == "":
        return jsonify({"error": "No file selected"}), 400

    if not allowed_file(file.filename):
        return jsonify({"error": "Only Excel files (.xlsx or .xls) are allowed"}), 400

    parent_run_id = generate_run_id()
    parent_folder = os.path.join(RUNS_FOLDER, parent_run_id)
    os.makedirs(parent_folder, exist_ok=True)

    upload_path = None

    try:
        upload_path = prepare_upload(file, parent_folder)
        df = read_input_dataframe(upload_path)

        batches = split_dataframe_for_dsw_batches(df)

        if not batches:
            shutil.rmtree(parent_folder, ignore_errors=True)
            return jsonify({"error": "No employees found for DSW batch processing"}), 400

        batch_summaries = []

        for batch in batches:
            batch_id = batch["id"]
            batch_folder = os.path.join(parent_folder, batch_id)
            os.makedirs(batch_folder, exist_ok=True)

            batch_input_path = os.path.join(batch_folder, "batch_input.xlsx")
            batch["df"].to_excel(batch_input_path, index=False)

            batch_summaries.append({
                "id": batch_id,
                "label": batch["label"],
                "count": batch["count"],
                "status": "waiting"
            })

        return jsonify({
            "run_id": parent_run_id,
            "batches": batch_summaries,
            "summary": {
                "total_employees": int(sum(b["count"] for b in batch_summaries)),
                "batch_count": len(batch_summaries)
            }
        })

    except ValueError as e:
        if os.path.exists(parent_folder):
            shutil.rmtree(parent_folder, ignore_errors=True)
        return jsonify({"error": str(e)}), 400

    except Exception:
        if os.path.exists(parent_folder):
            shutil.rmtree(parent_folder, ignore_errors=True)
        return jsonify({"error": "DSW batch setup failed"}), 500

    finally:
        if upload_path and os.path.exists(upload_path):
            try:
                os.remove(upload_path)
            except Exception:
                pass


@app.route("/api/run-adverse-batch/<parent_run_id>/<batch_id>", methods=["POST"])
def run_adverse_batch(parent_run_id, batch_id):
    auth_error = require_api_key()
    if auth_error:
        return auth_error

    if not re.fullmatch(r"batch-\d{3}", batch_id or ""):
        return jsonify({"error": "Invalid batch"}), 400

    parent_folder = os.path.join(RUNS_FOLDER, parent_run_id)
    batch_folder = os.path.join(parent_folder, batch_id)
    batch_input_path = os.path.join(batch_folder, "batch_input.xlsx")

    if not os.path.exists(batch_input_path):
        return jsonify({"error": "Batch input file not found"}), 404

    try:
        df = pd.read_excel(batch_input_path)
        employee_results = process_adverse_only_run(df, batch_folder)
        return make_batch_response(parent_run_id, batch_id, employee_results)

    except Exception:
        return jsonify({"error": f"DSW batch {batch_id} failed"}), 500


@app.route("/api/download/<parent_run_id>/<batch_id>/zip", methods=["GET"])
def download_batch_zip(parent_run_id, batch_id):
    auth_error = require_api_key()
    if auth_error:
        return auth_error

    if not re.fullmatch(r"batch-\d{3}", batch_id or ""):
        return jsonify({"error": "Invalid batch"}), 400

    parent_folder = os.path.join(RUNS_FOLDER, parent_run_id)
    batch_folder = os.path.join(parent_folder, batch_id)
    zip_path = None

    if not os.path.exists(batch_folder):
        return jsonify({"error": "Batch folder not found"}), 404

    for f in os.listdir(batch_folder):
        if f.endswith(".zip"):
            zip_path = os.path.join(batch_folder, f)
            break

    if not zip_path or not os.path.exists(zip_path):
        return jsonify({"error": "Batch ZIP file not found"}), 404

    return send_file(
        zip_path,
        as_attachment=True,
        download_name=f"DSW_Report_{batch_id}.zip"
    )


@app.route("/api/run-adverse", methods=["POST"])
def run_adverse_only():
    auth_error = require_api_key()
    if auth_error:
        return auth_error

    cleanup_old_runs(RUNS_FOLDER)
    cleanup_old_uploads(UPLOAD_FOLDER)

    file = request.files.get("file")
    if not file or file.filename == "":
        return jsonify({"error": "No file selected"}), 400

    if not allowed_file(file.filename):
        return jsonify({"error": "Only Excel files (.xlsx or .xls) are allowed"}), 400

    run_id = generate_run_id()
    run_folder = os.path.join(RUNS_FOLDER, run_id)
    os.makedirs(run_folder, exist_ok=True)

    upload_path = None

    try:
        upload_path = prepare_upload(file, run_folder)
        df = read_input_dataframe(upload_path)
        employee_results = process_adverse_only_run(df, run_folder)
        return make_response(run_id, employee_results)
    except ValueError as e:
        if os.path.exists(run_folder):
            shutil.rmtree(run_folder, ignore_errors=True)
        return jsonify({"error": str(e)}), 400
    except Exception:
        if os.path.exists(run_folder):
            shutil.rmtree(run_folder, ignore_errors=True)
        return jsonify({"error": "Processing failed"}), 500
    finally:
        if upload_path and os.path.exists(upload_path):
            try:
                os.remove(upload_path)
            except Exception:
                pass


@app.route("/api/download/<run_id>/single-pdf", methods=["GET"])
def download_single_pdf(run_id):
    auth_error = require_api_key()
    if auth_error:
        return auth_error

    run_folder = os.path.join(RUNS_FOLDER, run_id)
    pdf_path = None

    if not os.path.exists(run_folder):
        return jsonify({"error": "Run folder not found"}), 404

    for f in os.listdir(run_folder):
        if f.startswith("Single_Person_Check_") and f.endswith(".pdf"):
            pdf_path = os.path.join(run_folder, f)
            break

    if not pdf_path or not os.path.exists(pdf_path):
        return jsonify({"error": "PDF file not found"}), 404

    response = send_file(
        pdf_path,
        as_attachment=True,
        download_name=os.path.basename(pdf_path),
        mimetype="application/pdf"
    )

    def cleanup():
        try:
            shutil.rmtree(run_folder)
        except Exception:
            pass

    response.call_on_close(cleanup)
    return response


@app.route("/api/download/<run_id>/zip", methods=["GET"])
def download_zip(run_id):
    auth_error = require_api_key()
    if auth_error:
        return auth_error

    run_folder = os.path.join(RUNS_FOLDER, run_id)
    zip_path = None

    if not os.path.exists(run_folder):
        return jsonify({"error": "Run folder not found"}), 404

    for f in os.listdir(run_folder):
        if f.endswith(".zip"):
            zip_path = os.path.join(run_folder, f)
            break

    if not zip_path or not os.path.exists(zip_path):
        return jsonify({"error": "ZIP file not found"}), 404

    response = send_file(
        zip_path,
        as_attachment=True,
        download_name=os.path.basename(zip_path)
    )

    def cleanup():
        try:
            shutil.rmtree(run_folder)
        except Exception:
            pass

    response.call_on_close(cleanup)
    return response


@app.route("/api/download/<run_id>/results-excel", methods=["GET"])
def download_results_excel(run_id):
    auth_error = require_api_key()
    if auth_error:
        return auth_error

    return jsonify({"error": "Excel is included inside the ZIP file only"}), 410


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
