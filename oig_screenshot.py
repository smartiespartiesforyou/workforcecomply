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


def safe_part(value):
    value = str(value).strip()
    value = re.sub(r"[^A-Za-z0-9._-]+", "_", value)
    return value.strip("_") or "employee"


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
    except:
        pass

    _playwright = None
    _browser = None
    _context = None
    _page = None


atexit.register(close_oig_session)


def _wait_for_results(page):
    for _ in range(30):
        try:
            if page.locator("text=Search Results").first.is_visible(timeout=200):
                return True
            if page.locator("text=No Results").first.is_visible(timeout=200):
                return True
        except:
            pass
        page.wait_for_timeout(200)
    return False


def capture_oig(first_name, last_name, save_folder="proofs"):
    os.makedirs(save_folder, exist_ok=True)

    clean_first = safe_part(first_name)
    clean_last = safe_part(last_name)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")

    pdf_path = os.path.join(
        save_folder,
        f"OIG_{clean_first}_{clean_last}_{timestamp}.pdf"
    )

    page = _ensure_session()

    try:
        page.goto(OIG_URL, wait_until="domcontentloaded")

        page.fill(LAST_NAME_SELECTOR, last_name)
        page.fill(FIRST_NAME_SELECTOR, first_name)

        page.click(SEARCH_BUTTON_SELECTOR)

        _wait_for_results(page)

        page.pdf(
            path=pdf_path,
            format="Letter",
            print_background=True
        )

        return {"pdf_path": pdf_path}

    except Exception as e:
        return {"pdf_path": None, "error": str(e)}
