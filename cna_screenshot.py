import asyncio
import os
import re
from datetime import datetime
from threading import Thread
from playwright.async_api import async_playwright

CNA_URL = "https://tlc.dhh.la.gov/"
SSN_SELECTOR = "#txtSSNNum"
SEARCH_BUTTON_SELECTOR = "#btnSearch"


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


def build_cna_filename(first_name, last_name, ssn):
    first_clean = clean_name(first_name)
    last_clean = clean_name(last_name)
    last4 = get_last4(ssn)
    date_part = datetime.now().strftime("%Y-%m-%d")
    return f"{first_clean}_{last_clean}_{last4}_{date_part}_CNA.pdf"


def classify_cna_result(page_text):
    text = (page_text or "").lower()

    if "no records" in text or "no record" in text:
        return "not_found"

    if "expired" in text or "inactive" in text:
        return "not_active"

    if "active" in text:
        return "clear"

    if "certificate number" in text or "status" in text:
        return "review_needed"

    return "review_needed"


async def _wait_for_results(page):
    result_selectors = [
        "text=Certificate Number",
        "text=No records",
        "text=No Record",
        "text=Status",
    ]

    for _ in range(40):
        for selector in result_selectors:
            try:
                if await page.locator(selector).first.is_visible(timeout=250):
                    return True
            except Exception:
                pass
        await page.wait_for_timeout(250)

    return False


async def _open_cna_page(page):
    last_error = None

    for _ in range(3):
        try:
            await page.goto(CNA_URL, wait_until="commit", timeout=60000)
            await page.wait_for_selector(SSN_SELECTOR, timeout=60000)
            await page.wait_for_selector(SEARCH_BUTTON_SELECTOR, timeout=60000)
            return
        except Exception as e:
            last_error = e
            await page.wait_for_timeout(2000)

    raise last_error


async def _capture_cna_async(first_name, last_name, ssn, save_folder="proofs"):
    os.makedirs(save_folder, exist_ok=True)

    clean = clean_ssn(ssn)
    if not clean:
        return {
            "pdf_path": None,
            "error": "Invalid SSN",
            "cna_result": "invalid_ssn"
        }

    filename = build_cna_filename(first_name, last_name, ssn)
    pdf_path = os.path.join(save_folder, filename)

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
            await _open_cna_page(page)

            ssn_box = page.locator(SSN_SELECTOR)
            search_button = page.locator(SEARCH_BUTTON_SELECTOR)

            await ssn_box.fill("")
            await ssn_box.fill(clean)

            await search_button.click()

            await _wait_for_results(page)

            page_text = await page.content()
            cna_result = classify_cna_result(page_text)

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

            return {
                "pdf_path": pdf_path,
                "cna_result": cna_result
            }

        except Exception as e:
            return {
                "pdf_path": None,
                "error": str(e),
                "cna_result": "error"
            }

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


def capture_cna(first_name, last_name, ssn, save_folder="proofs"):
    return _run_async_in_thread(_capture_cna_async(first_name, last_name, ssn, save_folder))


def close_cna_session():
    return None
