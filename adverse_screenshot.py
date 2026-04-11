import asyncio
import os
import re
from datetime import datetime
from playwright.async_api import async_playwright

ADVERSE_URL = "https://adverseactions.ldh.la.gov/SelSearch"
SSN_SELECTOR = "#searchSsn"
SEARCH_BUTTON_SELECTOR = "#searchButton"
SEARCH_CONTAINER_SELECTOR = "#searchContainer"
DETAILS_SELECTOR = "#details"
LOADING_PANEL_SELECTOR = "#loadingPanel"

_browser = None
_context = None
_page = None


def clean_ssn(value):
    digits = re.sub(r"\D", "", str(value))
    if len(digits) != 9:
        return ""
    return f"{digits[:3]}-{digits[3:5]}-{digits[5:]}"


async def _init_browser():
    global _browser, _context, _page

    if _browser:
        return

    p = await async_playwright().start()

    _browser = await p.chromium.launch(
        executable_path="/usr/bin/chromium",
        headless=True,
        args=[
            "--no-sandbox",
            "--disable-dev-shm-usage",
        ],
    )

    _context = await _browser.new_context(
        viewport={"width": 1365, "height": 900},
        ignore_https_errors=True,
    )

    _page = await _context.new_page()
    _page.set_default_timeout(60000)


async def _wait_for_results(page):
    for _ in range(25):  # reduced loops
        try:
            try:
                if await page.locator(LOADING_PANEL_SELECTOR).is_visible(timeout=100):
                    await page.wait_for_timeout(200)
                    continue
            except:
                pass

            html = ""
            try:
                html = await page.content()
            except:
                pass

            text = html.lower()

            if "no results" in text or "<table" in text:
                return True

        except:
            pass

        await page.wait_for_timeout(200)

    return False


async def _capture_adverse_async(first_name, last_name, ssn="", save_folder="proofs"):
    global _page

    os.makedirs(save_folder, exist_ok=True)

    clean = clean_ssn(ssn)
    if not clean:
        return {"pdf_path": None, "error": "Invalid SSN"}

    await _init_browser()

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    pdf_path = os.path.join(save_folder, f"ADVERSE_{clean}_{timestamp}.pdf")

    try:
        await _page.goto(ADVERSE_URL, wait_until="domcontentloaded")

        ssn_box = _page.locator(SSN_SELECTOR)
        search_button = _page.locator(SEARCH_BUTTON_SELECTOR)

        await ssn_box.fill(clean)
        await search_button.click()

        await _wait_for_results(_page)

        await _page.pdf(
            path=pdf_path,
            format="Letter",
            print_background=True,
        )

        return {"pdf_path": pdf_path}

    except Exception as e:
        return {"pdf_path": None, "error": str(e)}


def capture_adverse(first_name, last_name, ssn="", save_folder="proofs"):
    return asyncio.run(_capture_adverse_async(first_name, last_name, ssn, save_folder))


def close_adverse_session():
    global _browser, _context, _page

    if _page:
        asyncio.run(_page.close())
    if _context:
        asyncio.run(_context.close())
    if _browser:
        asyncio.run(_browser.close())

    _browser = None
    _context = None
    _page = None
