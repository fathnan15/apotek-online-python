# playwright_runner.py

import re
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeoutError
from config import SIRS_APP_URL
from utils import reset_form

_playwright = None
_browser    = None
_page       = None

def init_sirs_manual(cdp_endpoint: str = "http://127.0.0.1:9222"):
    """
    Attach to your already-open Chrome profile.
    Then pause so you can manually choose filters and click 'Tampilkan'.
    """
    global _playwright, _browser, _page
    _playwright = sync_playwright().start()
    _browser    = _playwright.chromium.connect_over_cdp(cdp_endpoint)
    ctx         = _browser.contexts[0] if _browser.contexts else _browser.new_context()
    _page       = ctx.pages[0] if ctx.pages else ctx.new_page()

    # Navigate if you haven’t already:
    if _page.url != SIRS_APP_URL:
        try:
            _page.goto(SIRS_APP_URL, timeout=10000)
        except PWTimeoutError:
            print("⚠️  Timeout while navigating to SIRS app. Please ensure the URL is correct.")
            return
    # _page.goto(SIRS_APP_URL, timeout=10000)

    # urut = input("Enter the urut to filter by: ")
    # jenis_rawat = input("Enter the jenis_rawat to filter by: ")
    date = input("Enter the date (DD) to filter by: ")
    bulan = input("Enter the bulan to filter by: ")

    _page.select_option("#urut", "Nama")
    _page.select_option("#jenis_rawat", "Rawat Jalan")
    _page.select_option("#tanggal", date)
    _page.select_option("#bulan", bulan)
    _page.locator("input[type='button'][value='Tampilkan']").click()

    print("⏳ Waiting for process to start...")
    _page.wait_for_selector("#dv_process_start", state="attached", timeout=15000)
    print("⏳ Process started. Waiting for process to finish...")
    _page.wait_for_selector("#dv_process_start", state="detached", timeout=120000)
    print("✅ Process finished. Table should be visible.")
    
    # print("\n⚙️  Please switch to Chrome, set your filters, and click 'Tampilkan'.")
    # input("When the table is visible, press ⏎ Enter to continue…")


def get_claim_records() -> list[dict]:
    """
    After selecting filters manually or in test, scrape the JS-rendered table:
      - SEP from 4th <td>
      - receipt from Print Resep button onclick
    """
    if not _page:
        raise RuntimeError("Playwright page is not initialized. Call init_cdp() first.")
    rows = _page.locator("div#dv_content table.tblcontrast tbody tr")
    rows.first.wait_for(timeout=5000)
    records = []
    for i in range(rows.count()):
        row = rows.nth(i)
        sep_num     = row.locator("td").nth(3).inner_text().strip()
        mrn     = row.locator("td").nth(1).inner_text().strip().replace("-", "")
        dttm_sep     = row.locator("td").nth(4).inner_text().strip()
        onclick = row.locator("input[value='Print Resep']").get_attribute("onclick") or ""
        m = re.search(r'print_prescription\("([^"]+)"', onclick)
        receipt = m.group(1) if m else ""
        receipt = receipt[-5:] if receipt.isdigit() and len(receipt) >= 5 else receipt
        records.append({"dttm_sep":dttm_sep, "mrn" : mrn, "sep_num": sep_num, "receipt_num": receipt})

    _page.locator("input[type='button'][value='Download']").click()
    print("⏳ Downloading .....")
    return records

def close():
    """Tear down the SIRS Playwright session."""
    global _browser, _playwright
    if _browser:
        _browser.close()
    if _playwright:
        _playwright.stop()