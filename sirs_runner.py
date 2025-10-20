# playwright_runner.py

import re
import asyncio
from playwright.async_api import async_playwright, TimeoutError as PWTimeoutError
from config import SIRS_APP_URL
from utils import reset_form

_playwright = None
_browser    = None
_page       = None

async def init_sirs_manual(cdp_endpoint: str = "http://127.0.0.1:9222"):
    """
    Attach to your already-open Chrome profile.
    Then pause so you can manually choose filters and click 'Tampilkan'.
    """
    global _playwright, _browser, _page
    # REMOVE: async with async_playwright() as p:
    # The context is already set by set_playwright_context

    # Navigate if you haven’t already:
    if _page.url != SIRS_APP_URL:
        try:
            await _page.goto(SIRS_APP_URL, timeout=10000)
        except PWTimeoutError:
            print("⚠️  Timeout while navigating to SIRS app. Please ensure the URL is correct.")
            return

    date = input("Enter the date (DD) to filter by: ")
    bulan = input("Enter the bulan to filter by: ")

    await _page.select_option("#urut", "Nama")
    await _page.select_option("#jenis_rawat", "Rawat Jalan")
    await _page.select_option("#tanggal", date)
    await _page.select_option("#bulan", bulan)
    await _page.locator("input[type='button'][value='Tampilkan']").click()

    print("⏳ Waiting for process to start...")
    await _page.wait_for_selector("#dv_process_start", state="attached", timeout=15000)
    print("⏳ Process started. Waiting for process to finish...")
    await _page.wait_for_selector("#dv_process_start", state="detached", timeout=120000)
    print("✅ Process finished. Table should be visible.")
    
    # print("\n⚙️  Please switch to Chrome, set your filters, and click 'Tampilkan'.")
    # input("When the table is visible, press ⏎ Enter to continue…")


async def get_claim_records() -> list[dict]:
    """
    After selecting filters manually or in test, scrape the JS-rendered table:
      - SEP from 4th <td>
      - receipt from Print Resep button onclick
    """
    if not _page:
        raise RuntimeError("Playwright page is not initialized. Call init_cdp() first.")
    rows = _page.locator("div#dv_content table.tblcontrast tbody tr")
    await rows.first.wait_for(timeout=5000)
    count = await rows.count()
    records = []
    for i in range(count):
        row = rows.nth(i)
        sep_num     = await row.locator("td").nth(3).inner_text()
        mrn         = (await row.locator("td").nth(1).inner_text()).strip().replace("-", "")
        dttm_sep    = await row.locator("td").nth(4).inner_text()
        onclick     = await row.locator("input[value='Print Resep']").get_attribute("onclick") or ""
        m = re.search(r'print_prescription\("([^"]+)"', onclick)
        receipt = m.group(1) if m else ""
        receipt = receipt[-5:] if receipt.isdigit() and len(receipt) >= 5 else receipt
        records.append({"dttm_sep": dttm_sep, "mrn": mrn, "sep_num": sep_num, "receipt_num": receipt})
    return records

async def download_claims():
    if not _page:
        raise RuntimeError("Playwright page is not initialized. Call init_cdp() first.")
    print("⏳ Downloading .....", flush=True)

    async def handle_dialog(dialog):
        print(f"Dialog message: {dialog.message}", flush=True)
        await dialog.accept()

    _page.once("dialog", handle_dialog)  # Set handler before click

    await _page.locator("input[type='button'][value='Download']").click()
    try:
        # Optionally, wait for something that indicates download is complete
        await asyncio.sleep(2)  # Give time for dialog to appear and be handled
        return True
    except PWTimeoutError:
        return False

def close():
    """Tear down the SIRS Playwright session."""
    global _browser, _playwright
    if _browser:
        _browser.close()
    if _playwright:
        _playwright.stop()

async def set_playwright_context(p, cdp_endpoint: str = "http://127.0.0.1:9222"):
    global _playwright, _browser, _page
    _playwright = p
    _browser = await p.chromium.connect_over_cdp(cdp_endpoint)
    ctx = _browser.contexts[0] if _browser.contexts else await _browser.new_context()
    _page = ctx.pages[0] if ctx.pages else await ctx.new_page()