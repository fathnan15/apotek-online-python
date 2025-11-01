# apotek_runner.py

from playwright.sync_api import TimeoutError as PWTimeoutError, sync_playwright
from config import APOTEK_URL, APOTEK_SELECTORS

_playwright_apo = None
_browser_apo    = None
_page_apo       = None

def init_apotek(cdp_endpoint: str = "http://127.0.0.1:9222"):
    """Attach to Chrome CDP and navigate to the Apotek BPJS form."""
    global _playwright_apo, _browser_apo, _page_apo
    _playwright_apo = sync_playwright().start()
    _browser_apo    = _playwright_apo.chromium.connect_over_cdp(cdp_endpoint)
    ctx             = _browser_apo.contexts[0] if _browser_apo.contexts else _browser_apo.new_context()
    _page_apo       = ctx.pages[0] if ctx.pages else ctx.new_page()
    _page_apo.goto(APOTEK_URL, timeout=10000)
    print("✅ Connected to Apotek form.")

def submit_to_apotek(sep: str, receipt: str, rec_type: str) -> tuple[str, str]:
    sel = APOTEK_SELECTORS
    try:
        sep_str      = str(sep)
        receipt_str  = str(receipt)
        rec_type_str = str(rec_type)

        _page_apo.fill(sel['sep_input'], sep_str)
        _page_apo.type(sel['receipt_input'], receipt_str)
        _page_apo.click(sel['cari_button'])

        try:
            dlg = _page_apo.wait_for_event("dialog", timeout=1000)
            err_msg = dlg.message
            dlg.accept()
            _page_apo.click(sel['reset_button'])
            return ("error", err_msg)
        except PWTimeoutError:
            pass 

        _page_apo.wait_for_selector(
            sel['no_kartu_input'],
            state="attached",
            timeout=2000
        )

        _page_apo.fill(sel['receipt_type_input'], rec_type_str)
        _page_apo.wait_for_timeout(1000)

        _page_apo.click(sel['simpan_button'])

        try:
            dlg = _page_apo.wait_for_event("dialog", timeout=5000)
            msg = dlg.message
            dlg.accept()
            if "Simpan Berhasil" in msg:
                return ("normal", msg)
            _page_apo.click(sel['reset_button'])
            return ("error", msg)
        except PWTimeoutError:
            _page_apo.click(sel['reset_button'])
            return ("error", "No confirmation alert")

    except Exception as e:
        return ("error", str(e))
def close_apotek():
    """Tear down the Apotek Playwright session."""
    global _browser_apo, _playwright_apo
    if _browser_apo:
        _browser_apo.close()
    if _playwright_apo:
        _playwright_apo.stop()
