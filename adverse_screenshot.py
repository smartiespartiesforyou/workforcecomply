import asyncio
import os
import re
from datetime import datetime
from threading import Thread
from playwright.async_api import async_playwright

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


async def _capture_adverse_async(first_name, last_name, ssn="", save_folder="proofs"):
    os.makedirs(save_folder, exist_ok=True)

    clean = clean_ssn(ssn)
    if not clean:
        return {"pdf_path": None, "error": "Invalid SSN"}

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    pdf_path = os.path.join(save_folder, f"ADVERSE_{clean}_{timestamp}.pdf")

    async with async_playwright() as p:
        browser = await p.chromium.launch(
            executable_path="/usr/bin/chromium",
            headless=True,
            args=[
                "--no-sandbox",
                "--disable-dev-shm-usage",
                "--disable-blink-features=AutomationControlled",
            ],
        )

        context = await browser.new_context(
            viewport={"width": 1365, "height": 900},
            ignore_https_errors=True,
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            ),
        )

        page = await context.new_page()
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

            try:
                ssn_box = page.locator(SSN_SELECTOR)
                await ssn_box.evaluate(
                    "(el, value) => { el.value = value; el.setAttribute('value', value); }",
                    clean
                )
            except Exception:
                pass

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
            await context.close()
            await browser.close()


def _run_async_in_thread(coro):
    result = {}
    error = {}

    def runner():
        try:
            result["value"] = asyncio.run(coro)
        except Exception as e:
            error["value"] = e

    thread = Thread(target=runner)
    thread.start()
    thread.join()

    if "value" in error:
        raise error["value"]

    return result.get("value")


def capture_adverse(first_name, last_name, ssn="", save_folder="proofs"):
    return _run_async_in_thread(_capture_adverse_async(first_name, last_name, ssn, save_folder))


def close_adverse_session():
    return None
