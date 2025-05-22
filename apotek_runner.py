# apotek_runner.py

from playwright.sync_api import TimeoutError as PWTimeoutError, sync_playwright
from config import APOTEK_URL
from datetime import datetime

_playwright_apo = None
_browser_apo    = None
_page_apo       = None

def init_apotek(cdp_endpoint: str = "http://127.0.0.1:9222"):
    """Attach to your running Chrome and navigate to the Apotek BPJS form."""
    global _playwright_apo, _browser_apo, _page_apo
    _playwright_apo = sync_playwright().start()
    _browser_apo    = _playwright_apo.chromium.connect_over_cdp(cdp_endpoint)
    ctx             = _browser_apo.contexts[0] if _browser_apo.contexts else _browser_apo.new_context()
    _page_apo       = ctx.pages[0] if ctx.pages else ctx.new_page()
    _page_apo.goto(APOTEK_URL, timeout=10000)
    # If you need to log in first, do it here before returning.

def submit_to_apotek(sep: str, receipt: str, rec_type: str) -> tuple[str, str]:
    """
    1) Input SEP → click 'Cari'
    2) Wait for 'No Kartu' field to fill
    3) Select 'Jenis Resep'
    4) Click 'Simpan'
    5) Capture alert; on success return 'normal', on error click 'Reset' and return 'error'
    """
    global _page_apo
    try:
        # 1) Fill SEP and click Cari
        _page_apo.fill("input#TxtSEP", sep)               # ← adjust selector
        _page_apo.click("input[value='Cari']")            # ← adjust selector/value

        # 2) Wait for No Kartu to populate
        _page_apo.wait_for_function(
            "() => document.querySelector('input#TxtNoKartu').value.trim() !== ''",
            timeout=5000
        )

        # 3) Select the right 'Jenis Resep'
        _page_apo.select_option(
            "select#cboJnsResep",                         # ← adjust selector
            label=rec_type
        )

        # 4) Fill No Resep (if needed)
        _page_apo.fill("input#TxtNoResep", receipt)        # ← adjust selector

        # 5) Click Simpan
        _page_apo.click("input[value='Simpan']")           # ← adjust selector/value

        # 6) Handle the alert
        try:
            with _page_apo.expect_dialog(timeout=5000) as dlg:
                pass
            msg = dlg.value.message
            dlg.value.accept()

            if "Simpan Berhasil" in msg:
                return ("normal", msg)
            else:
                # on any other alert, reset form
                _page_apo.click("input[value='Reset']")    # ← adjust selector/value
                return ("error", msg)

        except PWTimeoutError:
            # No alert—treat as failure
            _page_apo.click("input[value='Reset']")        # ← adjust selector/value
            return ("error", "No confirmation alert")

    except Exception as e:
        # Unexpected exception
        return ("error", str(e))

def close_apotek():
    """Tear down the Apotek session."""
    global _browser_apo, _playwright_apo
    if _browser_apo:
        _browser_apo.close()
    if _playwright_apo:
        _playwright_apo.stop()
