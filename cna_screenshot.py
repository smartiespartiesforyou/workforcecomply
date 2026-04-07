import asyncio
import os
import re
from datetime import datetime
from playwright.async_api import async_playwright

CNA_URL = "https://tlc.dhh.la.gov/"
SSN_SELECTOR = "#txtSSNNum"
SEARCH_BUTTON_SELECTOR = "#btnSearch"


def clean_ssn(value):
    digits = re.sub(r"\D", "", str(value))
    if len(digits) != 9:
        return ""
    return f"{digits[:3]}-{digits[3:5]}-{digits[5:]}"


async def _wait_for_results(page):
    result_selectors = [
        "text=Certificate Number",
        "text=No records",
        "text=No Record",
        "text=Status",
    ]

    for _ in range(30):
        for selector in result_selectors:
            try:
                if await page.locator(selector).first.is_visible(timeout=200):
                    return True
            except Exception:
                pass
        await page.wait_for_timeout(200)

    return False


async def _capture_cna_async(ssn, save_folder="proofs"):
    os.makedirs(save_folder, exist_ok=True)

    clean = clean_ssn(ssn)
    if not clean:
        return {"pdf_path": None, "error": "Invalid SSN"}

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    pdf_path = os.path.join(save_folder, f"CNA_{clean}_{timestamp}.pdf")

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
        )

        page = await context.new_page()
        page.set_default_timeout(12000)
        page.set_default_navigation_timeout(20000)

        try:
            await page.goto(CNA_URL, wait_until="domcontentloaded")
            await page.wait_for_selector(SSN_SELECTOR)
            await page.wait_for_selector(SEARCH_BUTTON_SELECTOR)

            ssn_box = page.locator(SSN_SELECTOR)
            search_button = page.locator(SEARCH_BUTTON_SELECTOR)

            await ssn_box.fill("")
            await ssn_box.fill(clean)

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
            await context.close()
            await browser.close()


def capture_cna(ssn, save_folder="proofs"):
    return asyncio.run(_capture_cna_async(ssn, save_folder))


def close_cna_session():
    return None
