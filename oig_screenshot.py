from playwright.sync_api import sync_playwright
import os
import re
import atexit
from datetime import datetime

_playwright = None
_browser = None
_context = None
_page = None

OIG_URL = "https://exclusions.oig.hhs.gov/"
LAST_NAME_SELECTOR = "#ctl00_cpExclusions_txtSPLastName"
FIRST_NAME_SELECTOR = "#ctl00_cpExclusions_txtSPFirstName"
SEARCH_BUTTON_SELECTOR = "#ctl00_cpExclusions_ibSearchSP"


def clean_name(value):
    value = str(value).strip()
    value = re.sub(r"[^A-Za-z]", "", value)
    return value


def get_last4(ssn):
    digits = re.sub(r"\D", "", str(ssn))
    return digits[-4:] if len(digits) >= 4 else "0000"


def build_oig_filename(first_name, last_name, ssn):
    first_clean = clean_name(first_name)
    last_clean = clean_name(last_name)
    last4 = get_last4(ssn)
    date_part = datetime.now().strftime("%Y-%m-%d")
    return f"{first_clean}_{last_clean}_{last4}_{date_part}_OIG.pdf"


def _ensure_session():
    global _playwright, _browser, _context, _page

    if _playwright is None:
        _playwright = sync_playwright().start()

    if _browser is None:
        _browser = _playwright.chromium.launch(
            executable_path="/usr/bin/chromium",
            headless=True,
            args=[
                "--no-sandbox",
                "--disable-dev-shm-usage",
                "--disable-blink-features=AutomationControlled",
            ],
        )

    if _context is None:
        _context = _browser.new_context(
            viewport={"width": 1365, "height": 900},
        )

    if _page is None:
        _page = _context.new_page()
        _page.set_default_timeout(12000)

    return _page


def close_oig_session():
    global _playwright, _browser, _context, _page

    try:
        if _page:
            _page.close()
        if _context:
            _context.close()
        if _browser:
            _browser.close()
        if _playwright:
            _playwright.stop()
    except Exception:
        pass

    _playwright = None
    _browser = None
    _context = None
    _page = None


atexit.register(close_oig_session)


def _wait_for_results(page):
    for _ in range(40):
        try:
            body_text = page.locator("body").inner_text(timeout=1000).upper()

            if "NO RESULTS" in body_text:
                return "no_results"

            if "SEARCH RESULTS" in body_text:
                return "search_results"

        except Exception:
            pass

        page.wait_for_timeout(250)

    return "unknown"


def _detect_oig_match(page):
    try:
        body_text = page.locator("body").inner_text(timeout=5000).upper()

        if "NO RESULTS" in body_text:
            return False, "clear"

        if "SEARCH RESULTS" in body_text and "NO RESULTS" not in body_text:
            return True, "review_needed"

        return False, "error"

    except Exception:
        return False, "error"


def capture_oig(first_name, last_name, ssn, save_folder="proofs"):
    os.makedirs(save_folder, exist_ok=True)

    filename = build_oig_filename(first_name, last_name, ssn)
    pdf_path = os.path.join(save_folder, filename)

    page = _ensure_session()

    try:
        page.goto(OIG_URL, wait_until="domcontentloaded")

        page.fill(LAST_NAME_SELECTOR, "")
        page.fill(FIRST_NAME_SELECTOR, "")

        page.fill(LAST_NAME_SELECTOR, last_name)
        page.fill(FIRST_NAME_SELECTOR, first_name)

        page.click(SEARCH_BUTTON_SELECTOR)

        _wait_for_results(page)

        oig_match_found, oig_status = _detect_oig_match(page)

        page.pdf(
            path=pdf_path,
            format="Letter",
            print_background=True
        )

        return {
            "pdf_path": pdf_path,
            "oig_match_found": oig_match_found,
            "oig_status": oig_status
        }

    except Exception as e:
        return {
            "pdf_path": None,
            "oig_match_found": False,
            "oig_status": "error",
            "error": str(e)
        }
