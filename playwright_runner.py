# playwright_runner.py

import re
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeoutError
from config import SIRS_APP_URL
from utils import reset_form

_playwright = None
_browser    = None
_page       = None

def init_cdp(cdp_endpoint: str = "http://127.0.0.1:9222"):
    """Attach to your manually started Chrome and navigate to SIRS."""
    global _playwright, _browser, _page
    _playwright = sync_playwright().start()
    _browser    = _playwright.chromium.connect_over_cdp(cdp_endpoint)
    ctx         = _browser.contexts[0] if _browser.contexts else _browser.new_context()
    _page       = ctx.pages[0] if ctx.pages else ctx.new_page()
    _page.goto(SIRS_APP_URL, timeout=10000)

def get_claim_records() -> list[dict]:
    """
    After selecting filters manually or in test, scrape the JS-rendered table:
      - SEP from 4th <td>
      - receipt from Print Resep button onclick
    """
    rows = _page.locator("div#dv_content table.tblcontrast tbody tr")
    rows.first.wait_for(timeout=5000)
    records = []
    for i in range(rows.count()):
        row = rows.nth(i)
        sep     = row.locator("td").nth(3).inner_text().strip()
        onclick = row.locator("input[value='Print Resep']").get_attribute("onclick") or ""
        m = re.search(r'print_prescription\("([^"]+)"', onclick)
        receipt = m.group(1) if m else ""
        records.append({"sep": sep, "receipt": receipt})
    return records

def close():
    """Tear down the SIRS Playwright session."""
    global _browser, _playwright
    if _browser:
        _browser.close()
    if _playwright:
        _playwright.stop()
