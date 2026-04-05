from playwright.sync_api import sync_playwright
import os
import re
import atexit
from datetime import datetime

_playwright = None
_browser = None
_context = None
_page = None

ADVERSE_URL = "https://adverseactions.ldh.la.gov/SelSearch"
SSN_SELECTOR = "#searchSsn"
SEARCH_BUTTON_SELECTOR = "#searchButton"
SEARCH_CONTAINER_SELECTOR = "#searchContainer"
DETAILS_SELECTOR = "#details"
LOADING_PANEL_SELECTOR = "#loadingPanel"


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
            headless=True,
            args=[
                "--disable-blink-features=AutomationControlled",
                "--disable-dev-shm-usage",
                "--no-first-run",
                "--no-default-browser-check",
            ],
        )

    if _context is None:
        _context = _browser.new_context(viewport={"width": 1365, "height": 900})

    if _page is None:
        _page = _context.new_page()
        _page.set_default_timeout(12000)
        _page.set_default_navigation_timeout(15000)
        _page.goto(ADVERSE_URL, wait_until="domcontentloaded")
        _page.wait_for_selector(SSN_SELECTOR)
        _page.wait_for_selector(SEARCH_BUTTON_SELECTOR)

    return _page


def close_adverse_session():
    global _playwright, _browser, _context, _page

    try:
        if _page:
            _page.close()
    except Exception:
        pass

    try:
        if _context:
            _context.close()
    except Exception:
        pass

    try:
        if _browser:
            _browser.close()
    except Exception:
        pass

    try:
        if _playwright:
            _playwright.stop()
    except Exception:
        pass

    _playwright = None
    _browser = None
    _context = None
    _page = None


atexit.register(close_adverse_session)


def _ready_search_page(page):
    if "adverseactions.ldh.la.gov/selsearch" not in page.url.lower():
        page.goto(ADVERSE_URL, wait_until="domcontentloaded")

    page.wait_for_selector(SSN_SELECTOR)
    page.wait_for_selector(SEARCH_BUTTON_SELECTOR)


def _wait_for_results(page):
    for _ in range(32):
        try:
            loading_visible = False
            try:
                loading_visible = page.locator(LOADING_PANEL_SELECTOR).is_visible(timeout=100)
            except Exception:
                loading_visible = False

            if loading_visible:
                page.wait_for_timeout(200)
                continue

            search_html = ""
            details_html = ""

            try:
                search_html = page.locator(SEARCH_CONTAINER_SELECTOR).inner_html(timeout=100)
            except Exception:
                pass

            try:
                details_html = page.locator(DETAILS_SELECTOR).inner_html(timeout=100)
            except Exception:
                pass

            combined = f"{search_html} {details_html}".lower()

            if "no results" in combined:
                return True
            if "<table" in combined:
                return True
            if "first name" in combined and "last name" in combined:
                return True
            if "social security number" in combined and "search" not in combined:
                return True

        except Exception:
            pass

        page.wait_for_timeout(200)

    return False


def capture_adverse(first_name, last_name, ssn="", save_folder="proofs"):
    os.makedirs(save_folder, exist_ok=True)

    clean = clean_ssn(ssn)
    if not clean:
        print("ADVERSE ERROR: Invalid SSN")
        return {"pdf_path": None, "error": "Invalid SSN"}

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")

    pdf_path = os.path.join(
        save_folder,
        f"ADVERSE_{clean}_{timestamp}.pdf"
    )

    page = _ensure_session()

    try:
        _ready_search_page(page)

        ssn_box = page.locator(SSN_SELECTOR)
        search_button = page.locator(SEARCH_BUTTON_SELECTOR)

        ssn_box.fill("")
        ssn_box.fill(clean)

        ssn_box.evaluate(
            "(el, value) => { el.value = value; el.setAttribute('value', value); }",
            clean
        )

        search_button.click()

        _wait_for_results(page)

        try:
            ssn_box = page.locator(SSN_SELECTOR)
            ssn_box.evaluate(
                "(el, value) => { el.value = value; el.setAttribute('value', value); }",
                clean
            )
        except Exception:
            pass

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

        print("Saved ADVERSE PDF:", pdf_path)

        return {"pdf_path": pdf_path}

    except Exception as e:
        print("ADVERSE ERROR:", e)
        return {"pdf_path": None, "error": str(e)}
