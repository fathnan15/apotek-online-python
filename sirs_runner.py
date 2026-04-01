# playwright_runner.py

import re
import asyncio
from playwright.async_api import async_playwright, TimeoutError as PWTimeoutError
from config import SIRS_APP_URLS
from utils import reset_form

_playwright = None
_browser    = None
_page       = None

async def init_sirs_manual(cdp_endpoint: str = "http://127.0.0.1:9222", date: str | None = None, bulan: str | None = None, server: str = "223"):
    """
    Attach to your already-open Chrome profile and navigate to the assigned server.
    """
    global _playwright, _browser, _page

    target_url = SIRS_APP_URLS[server]

    # Navigate if you haven’t already:
    if _page.url != target_url:
        try:
            await _page.goto(target_url, timeout=10000)
        except PWTimeoutError:
            print(f"⚠️  Timeout while navigating to SIRS app at {target_url}.")
            return

    # prompt only when not provided
    if date is None:
        date = input("Enter the date (DD) to filter by: ")
    if bulan is None:
        bulan = input("Enter the bulan to filter by: ")

    await _page.select_option("#urut", "Nama")
    await _page.select_option("#jenis_rawat", "Rawat Jalan")
    await _page.select_option("#tanggal", date)
    await _page.select_option("#bulan", bulan)

    # 1. Wait for any previous overlay to clear
    try:
        await _page.locator("#dv_process_start").wait_for(state="hidden", timeout=60000)
    except PWTimeoutError:
        pass

    # 2. Click the button to request data
    await _page.locator("input[type='button'][value='Tampilkan']").click()
    
    # 3. CRITICAL NEW STEP: Wait for the new overlay triggered by Tampilkan to finish
    try:
        await _page.locator("#dv_process_start").wait_for(state="hidden", timeout=90000) # 90s safety margin
    except PWTimeoutError:
        print("⚠️ Warning: Data query took longer than 90s. Server may be unresponsive.")
    
    # print("\n⚙️  Please switch to Chrome, set your filters, and click 'Tampilkan'.")
    # input("When the table is visible, press ⏎ Enter to continue…")


# async def get_claim_records() -> list[dict]:
#     """
#     After selecting filters manually or in test, scrape the JS-rendered table:
#       - SEP from 4th <td>
#       - receipt from Print Resep button onclick
#     """
#     if not _page:
#         raise RuntimeError("Playwright page is not initialized. Call init_cdp() first.")
#     rows = _page.locator("div#dv_content table.tblcontrast tbody tr")
#     await rows.first.wait_for(timeout=5000)
#     count = await rows.count()
#     records = []
#     for i in range(count):
#         row = rows.nth(i)
#         sep_num     = await row.locator("td").nth(3).inner_text()
#         mrn         = (await row.locator("td").nth(1).inner_text()).strip().replace("-", "")
#         dttm_sep    = await row.locator("td").nth(4).inner_text()
#         onclick     = await row.locator("input[value='Print Resep']").get_attribute("onclick") or ""
#         m = re.search(r'print_prescription\("([^"]+)"', onclick)
#         receipt = m.group(1) if m else ""
#         receipt = receipt[-5:] if receipt.isdigit() and len(receipt) >= 5 else receipt
#         records.append({"dttm_sep": dttm_sep, "mrn": mrn, "sep_num": sep_num, "receipt_num": receipt})
#     return records

async def get_claim_records() -> list[dict]:
    """
    Scrapes the JS-rendered table with retry logic to handle slow rendering.
    """
    if not _page:
        raise RuntimeError("Playwright page is not initialized. Call init_cdp() first.")

    # 1. Wait for the table container to be visible
    container = _page.locator("div#dv_content table.tblcontrast")
    await container.wait_for(state="visible", timeout=60000)

    # 2. Retry logic for row population
    # BPJS tables sometimes show 0 rows initially even after the loading spinner hides
    records = []
    max_retries = 5
    for attempt in range(max_retries):
        rows = container.locator("tbody tr")
        count = await rows.count()
        
        if count > 0:
            print(f"📊 Found {count} records. Starting extraction...")
            for i in range(count):
                row = rows.nth(i)
                # Verify row has actual data (td index 3 for SEP)
                sep_num = await row.locator("td").nth(3).inner_text()
                if not sep_num.strip():
                    continue
                    
                mrn = (await row.locator("td").nth(1).inner_text()).strip().replace("-", "")
                dttm_sep = await row.locator("td").nth(4).inner_text()
                
                # Robust extraction for the receipt number
                onclick = await row.locator("input[value='Print Resep']").get_attribute("onclick") or ""
                m = re.search(r'print_prescription\("([^"]+)"', onclick)
                presc_id = (m.group(1) or "") if m else ""
                receipt = "2" + presc_id[-4:] if presc_id.isdigit() and len(presc_id) >= 4 else presc_id
                
                records.append({
                    "dttm_sep": dttm_sep, 
                    "mrn": mrn, 
                    "sep_num": sep_num, 
                    "presc_id" : presc_id,
                    "receipt_num": receipt
                })
            return records
        
        print(f"⚠️ Table empty (Attempt {attempt+1}/{max_retries}). Retrying in 2s...")
        await asyncio.sleep(2)

    print("❌ Failed to find any records in the table after multiple attempts.")
    return []
    
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
        # Wait for download overlay to clear before yielding back to the main loop
        await _page.locator("#dv_process_start").wait_for(state="hidden", timeout=60000)
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