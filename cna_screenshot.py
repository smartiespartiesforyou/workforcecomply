from playwright.sync_api import sync_playwright
import os
import re
import atexit
from datetime import datetime

_playwright = None
_browser = None
_context = None
_page = None

CNA_URL = "https://tlc.dhh.la.gov/"
SSN_SELECTOR = "#txtSSNNum"
SEARCH_BUTTON_SELECTOR = "#btnSearch"


def clean_ssn(value):
    digits = re.sub(r"\D", "", str(value))
    if len(digits) != 9:
        return ""
    return f"{digits[:3]}-{digits[3:5]}-{digits[5:]}"


def _ensure_session():
    global _playwright, _browser, _context, _page

    if _playwright is None:
        _playwright = sync_playwright().start()

    if _browser is None:
        _browser = _playwright.chromium.launch(headless=True)

    if _context is None:
        _context = _browser.new_context(viewport={"width": 1365, "height": 900})

    if _page is None:
        _page = _context.new_page()

    return _page


def close_cna_session():
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
    except:
        pass

    _playwright = None
    _browser = None
    _context = None
    _page = None


atexit.register(close_cna_session)


def capture_cna(ssn, save_folder="proofs"):
    os.makedirs(save_folder, exist_ok=True)

    clean = clean_ssn(ssn)
    if not clean:
        print("CNA ERROR: Invalid SSN")
        return {"pdf_path": None, "error": "Invalid SSN"}

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")

    pdf_path = os.path.join(
        save_folder,
        f"CNA_{clean}_{timestamp}.pdf"
    )

    page = _ensure_session()

    try:
        page.goto(CNA_URL, wait_until="domcontentloaded")

        page.wait_for_selector(SSN_SELECTOR)
        page.wait_for_selector(SEARCH_BUTTON_SELECTOR)

        ssn_box = page.locator(SSN_SELECTOR)
        search_button = page.locator(SEARCH_BUTTON_SELECTOR)

        ssn_box.fill("")
        ssn_box.fill(clean)

        search_button.click()

        for _ in range(30):
            try:
                if page.locator("text=Certificate Number").first.is_visible(timeout=200):
                    break
                if page.locator("text=No records").first.is_visible(timeout=200):
                    break
                if page.locator("text=No Record").first.is_visible(timeout=200):
                    break
                if page.locator("text=Status").first.is_visible(timeout=200):
                    break
            except:
                pass
            page.wait_for_timeout(200)

        page.pdf(
            path=pdf_path,
            format="Letter",
            print_background=True
        )

        print("Saved CNA PDF:", pdf_path)

        return {"pdf_path": pdf_path}

    except Exception as e:
        print("CNA ERROR:", e)
        return {"pdf_path": None, "error": str(e)}
