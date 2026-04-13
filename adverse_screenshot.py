import asyncio
import os
import re
from datetime import datetime
from threading import Thread, Lock
from playwright.async_api import async_playwright

ADVERSE_URL = "https://adverseactions.ldh.la.gov/SelSearch"
SSN_SELECTOR = "#searchSsn"
SEARCH_BUTTON_SELECTOR = "#searchButton"
SEARCH_CONTAINER_SELECTOR = "#searchContainer"
DETAILS_SELECTOR = "#details"
LOADING_PANEL_SELECTOR = "#loadingPanel"

_worker_thread = None
_worker_loop = None
_worker_lock = Lock()

_playwright = None
_browser = None
_context = None


def clean_ssn(value):
    digits = re.sub(r"\D", "", str(value))
    if len(digits) != 9:
        return ""
    return f"{digits[:3]}-{digits[3:5]}-{digits[5:]}"


def clean_name(value):
    value = str(value).strip()
    value = re.sub(r"[^A-Za-z]", "", value)
    return value


def get_last4(ssn):
    digits = re.sub(r"\D", "", str(ssn))
    return digits[-4:] if len(digits) >= 4 else "0000"


def build_dsw_filename(first_name, last_name, ssn):
    first_clean = clean_name(first_name)
    last_clean = clean_name(last_name)
    last4 = get_last4(ssn)
    date_part = datetime.now().strftime("%Y-%m-%d")

    return f"{first_clean}_{last_clean}_{last4}_{date_part}_DSW.pdf"


async def _wait_for_results(page):
    for _ in range(40):
        try:
            loading_visible = False

            try:
                loading_visible = await page.locator(LOADING_PANEL_SELECTOR).is_visible(timeout=150)
            except Exception:
                loading_visible = False

            if loading_visible:
                await page.wait_for_timeout(250)
                continue

            search_html = ""
            details_html = ""

            try:
                search_html = await page.locator(SEARCH_CONTAINER_SELECTOR).inner_html(timeout=150)
            except Exception:
                pass

            try:
                details_html = await page.locator(DETAILS_SELECTOR).inner_html(timeout=150)
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

        await page.wait_for_timeout(250)

    return False


async def _open_adverse_page(page):
    last_error = None

    for _ in range(3):
        try:
            await page.goto(ADVERSE_URL, wait_until="commit", timeout=60000)
            await page.wait_for_selector(SSN_SELECTOR, timeout=60000)
            await page.wait_for_selector(SEARCH_BUTTON_SELECTOR, timeout=60000)
            return
        except Exception as e:
            last_error = e
            await page.wait_for_timeout(2000)

    raise last_error


def _start_worker_loop():
    global _worker_loop

    loop = asyncio.new_event_loop()
    _worker_loop = loop
    asyncio.set_event_loop(loop)
    loop.run_forever()


def _ensure_worker():
    global _worker_thread

    with _worker_lock:
        if _worker_thread and _worker_thread.is_alive() and _worker_loop:
            return

        _worker_thread = Thread(target=_start_worker_loop, daemon=True)
        _worker_thread.start()

        while _worker_loop is None:
            pass


async def _ensure_browser():
    global _playwright, _browser, _context

    if _browser is not None and _context is not None:
        return

    _playwright = await async_playwright().start()

    _browser = await _playwright.chromium.launch(
        executable_path="/usr/bin/chromium",
        headless=True,
        args=[
            "--no-sandbox",
            "--disable-dev-shm-usage",
            "--disable-blink-features=AutomationControlled",
        ],
    )

    _context = await _browser.new_context(
        viewport={"width": 1365, "height": 900},
        ignore_https_errors=True,
        user_agent=(
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/124.0.0.0 Safari/537.36"
        ),
    )


async def _capture_adverse_async(first_name, last_name, ssn="", save_folder="proofs"):
    global _context

    os.makedirs(save_folder, exist_ok=True)

    clean = clean_ssn(ssn)
    if not clean:
        return {"pdf_path": None, "error": "Invalid SSN"}

    filename = build_dsw_filename(first_name, last_name, ssn)
    pdf_path = os.path.join(save_folder, filename)

    await _ensure_browser()

    page = await _context.new_page()
    page.set_default_timeout(60000)
    page.set_default_navigation_timeout(60000)

    try:
        await _open_adverse_page(page)

        ssn_box = page.locator(SSN_SELECTOR)
        search_button = page.locator(SEARCH_BUTTON_SELECTOR)

        await ssn_box.fill("")
        await ssn_box.fill(clean)

        await ssn_box.evaluate(
            "(el, value) => { el.value = value; el.setAttribute('value', value); }",
            clean
        )

        await search_button.click()

        await _wait_for_results(page)

        await page.emulate_media(media="screen")
        await page.pdf(
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

    finally:
        await page.close()


def capture_adverse(first_name, last_name, ssn="", save_folder="proofs"):
    _ensure_worker()
    future = asyncio.run_coroutine_threadsafe(
        _capture_adverse_async(first_name, last_name, ssn, save_folder),
        _worker_loop,
    )
    return future.result()


async def _shutdown_async():
    global _playwright, _browser, _context

    try:
        if _context is not None:
            await _context.close()
    except Exception:
        pass

    try:
        if _browser is not None:
            await _browser.close()
    except Exception:
        pass

    try:
        if _playwright is not None:
            await _playwright.stop()
    except Exception:
        pass

    _context = None
    _browser = None
    _playwright = None


def close_adverse_session():
    global _worker_thread, _worker_loop

    if _worker_loop is None:
        return None

    try:
        future = asyncio.run_coroutine_threadsafe(_shutdown_async(), _worker_loop)
        future.result(timeout=30)
    except Exception:
        pass

    try:
        _worker_loop.call_soon_threadsafe(_worker_loop.stop)
    except Exception:
        pass

    if _worker_thread and _worker_thread.is_alive():
        _worker_thread.join(timeout=5)

    _worker_thread = None
    _worker_loop = None
    return None
