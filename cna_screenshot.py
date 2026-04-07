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
        _page.set_default_navigation_timeout(20000)

    return _page


def close_cna_session():
    global _playwright, _browser, _context, _page

    try:
        if _page is not None:
            _page.close()
    except Exception:
        pass

    try:
        if _context is not None:
            _context.close()
    except Exception:
        pass

    try:
        if _browser is not None:
            _browser.close()
    except Exception:
        pass

    try:
        if _playwright is not None:
            _playwright.stop()
    except Exception:
        pass

    _page = None
    _context = None
    _browser = None
    _playwright = None


atexit.register(close_cna_session)


def _go_to_search_page(page):
    page.goto(CNA_URL, wait_until="domcontentloaded")
    page.wait_for_selector(SSN_SELECTOR)
    page.wait_for_selector(SEARCH_BUTTON_SELECTOR)


def _wait_for_results(page):
    result_selectors = [
        "text=Certificate Number",
        "text=No records",
        "text=No Record",
        "text=Status",
    ]

    for _ in range(30):
        for selector in result_selectors:
            try:
                if page.locator(selector).first.is_visible(timeout=200):
                    return True
            except Exception:
                pass
        page.wait_for_timeout(200)

    return False


def capture_cna(ssn, save_folder="proofs"):
    os.makedirs(save_folder, exist_ok=True)

    clean = clean_ssn(ssn)
    if not clean:
        return {"pdf_path": None, "error": "Invalid SSN"}

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")

    pdf_path = os.path.join(
        save_folder,
        f"CNA_{clean}_{timestamp}.pdf"
    )

    page = _ensure_session()

    try:
        _go_to_search_page(page)

        ssn_box = page.locator(SSN_SELECTOR)
        search_button = page.locator(SEARCH_BUTTON_SELECTOR)

        ssn_box.fill("")
        ssn_box.fill(clean)

        search_button.click()

        _wait_for_results(page)

        page.emulate_media(media="screen")
        page.pdf(
            path=pdf_path,
            format="Letter",
            print_background=True,
            margin={
                "top": "0.2in",
                "right": "0.2in",
                "bottom": "0.2in",
                "left": "0.2in",
            },
        )

        return {"pdf_path": pdf_path}

    except Exception as e:
        return {"pdf_path": None, "error": str(e)}
