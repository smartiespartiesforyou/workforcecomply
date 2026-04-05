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
            headless=True,
            args=[
                "--disable-blink-features=AutomationControlled",
                "--disable-dev-shm-usage",
                "--no-first-run",
                "--no-default-browser-check",
            ],
        )

    if _context is None:
        _context = _browser.new_context(
            viewport={"width": 1365, "height": 900},
            accept_downloads=False,
        )

    if _page is None:
        _page = _context.new_page()
        _page.set_default_timeout(12000)
        _page.set_default_navigation_timeout(20000)

    _go_to_search_page(_page)
    return _page


def _go_to_search_page(page):
    page.goto(OIG_URL, wait_until="domcontentloaded")
    page.wait_for_selector(LAST_NAME_SELECTOR)
    page.wait_for_selector(FIRST_NAME_SELECTOR)
    page.wait_for_selector(SEARCH_BUTTON_SELECTOR)


def close_oig_session():
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


atexit.register(close_oig_session)


def _wait_for_results(page):
    result_selectors = [
        "#ctl00_cpExclusions_gvSearchSP",
        "#ctl00_cpExclusions_lblSPSearchResults",
        "#ctl00_cpExclusions_pnlSearchSPResults",
        "text=No Results Were Found",
        "text=No results were found",
        "text=Search Results",
    ]

    for _ in range(32):
        for selector in result_selectors:
            try:
                if page.locator(selector).first.is_visible(timeout=250):
                    return True
            except Exception:
                pass
        page.wait_for_timeout(250)

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
        # Always start each employee from a clean search page.
        _go_to_search_page(page)

        last_box = page.locator(LAST_NAME_SELECTOR)
        first_box = page.locator(FIRST_NAME_SELECTOR)
        search_button = page.locator(SEARCH_BUTTON_SELECTOR)

        last_box.fill("")
        first_box.fill("")

        last_box.fill(str(last_name).strip())
        first_box.fill(str(first_name).strip())

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

        print("Saved PDF:", pdf_path)

        return {
            "pdf_path": pdf_path,
        }

    except Exception as e:
        print("ERROR:", e)
        return {
            "pdf_path": None,
            "error": str(e),
        }


if __name__ == "__main__":
    try:
        capture_oig("Floyd", "Holmes")
    finally:
        close_oig_session()
