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
    _page.goto(SIRS_APP_URL, timeout=10000)

    print("\n⚙️  Please switch to Chrome, set your filters, and click 'Tampilkan'.")
    input("When the table is visible, press ⏎ Enter to continue…")


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
        dttm_sep     = row.locator("td").nth(4).inner_text().strip()
        onclick = row.locator("input[value='Print Resep']").get_attribute("onclick") or ""
        m = re.search(r'print_prescription\("([^"]+)"', onclick)
        receipt = m.group(1) if m else ""
        receipt = receipt[-5:] if receipt.isdigit() and len(receipt) >= 5 else receipt
        records.append({"dttm_sep":dttm_sep, "sep_num": sep_num, "receipt_num": receipt})
    return records

def close():
    """Tear down the SIRS Playwright session."""
    global _browser, _playwright
    if _browser:
        _browser.close()
    if _playwright:
        _playwright.stop()