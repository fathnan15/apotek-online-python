# apotek_runner.py

from playwright.sync_api import TimeoutError as PWTimeoutError, sync_playwright
from config import APOTEK_URL, APOTEK_SELECTORS
import time

_playwright_apo = None
_browser_apo    = None
_page_apo       = None

def _now_ms():
    return int(time.time() * 1000)

def _adaptive_wait_for_function(page, js_func, arg, fast_timeout=800, long_timeout=7000, poll_interval=0.08):
    """
    Try a short fast_timeout first; if not satisfied, extend to long_timeout.
    Returns True if the function returned truthy within total time, False otherwise.
    Uses page.evaluate in a tight loop (less overhead than repeated full Playwright waits).
    `js_func` should be a JS snippet that returns a value when done; we'll call it via evaluate.
    We call evaluate repeatedly with a small sleep to keep responsiveness.
    """
    start = _now_ms()
    deadline_fast = start + fast_timeout
    deadline_long = start + long_timeout

    # First, fast loop
    while _now_ms() <= deadline_fast:
        try:
            val = page.evaluate(js_func, arg)
            if val:
                return True
        except Exception:
            # transient page state — ignore and retry
            pass
        time.sleep(poll_interval)

    # Second, longer loop
    while _now_ms() <= deadline_long:
        try:
            val = page.evaluate(js_func, arg)
            if val:
                return True
        except Exception:
            pass
        time.sleep(poll_interval)

    return False


def init_apotek(cdp_endpoint: str = "http://127.0.0.1:9222"):
    """Attach to Chrome CDP and navigate to the Apotek BPJS form."""
    global _playwright_apo, _browser_apo, _page_apo
    _playwright_apo = sync_playwright().start()
    _browser_apo    = _playwright_apo.chromium.connect_over_cdp(cdp_endpoint)
    ctx             = _browser_apo.contexts[0] if _browser_apo.contexts else _browser_apo.new_context()
    _page_apo       = ctx.pages[0] if ctx.pages else ctx.new_page()
    # keep default timeout reasonably small — adaptive waits handle slow cases
    _page_apo.set_default_timeout(4000)  # 4s
    _page_apo.goto(APOTEK_URL, timeout=10000)
    print("✅ Connected to Apotek form.")


def submit_to_apotek(sep: str, receipt: str, rec_type: str) -> tuple[str, str]:
    sel = APOTEK_SELECTORS
    try:
        sep_str      = str(sep)
        receipt_str  = str(receipt)
        rec_type_str = str(rec_type)

        # fill SEP and trigger search (keeps it fast)
        _page_apo.fill(sel['sep_input'], sep_str)
        _page_apo.keyboard.press("Enter")

        # Wait adaptively for the card-number input to hold a non-empty value.
        js_check_value = """(selector) => {
            const el = document.querySelector(selector);
            if (!el) return false;
            const v = el.value;
            return v !== null && v !== undefined && v.toString().trim().length > 0;
        }"""

        ok = _adaptive_wait_for_function(
            _page_apo,
            js_check_value,
            sel['no_kartu_input'],
            fast_timeout=200,    # quick path: 0.7s
            long_timeout=700,   # slow path: up to 7s
            poll_interval=0.06
        )

        if not ok:
            return ("error", "No card number returned by page")

        # It's common that an immediate dialog (error) appears after search;
        # try a very short wait first, then a slightly longer one if needed.
        try:
            dlg = _page_apo.wait_for_event("dialog", timeout=700)
            err_msg = dlg.message
            dlg.accept()
            _page_apo.click(sel['reset_button'])
            return ("error", err_msg)
        except PWTimeoutError:
            # short wait didn't find dialog; try a longer but still bounded wait
            try:
                dlg = _page_apo.wait_for_event("dialog", timeout=700)
                err_msg = dlg.message
                dlg.accept()
                _page_apo.click(sel['reset_button'])
                return ("error", err_msg)
            except PWTimeoutError:
                pass

        # fill receipt type and receipt number (fill is faster than type with delay)
        _page_apo.fill(sel['receipt_type_input'], rec_type_str)
        _page_apo.fill(sel['receipt_input'], receipt_str)
        # tiny pause to let the page process the filled value (very short)
        _page_apo.wait_for_timeout(120)
        _page_apo.click(sel['simpan_button'])

        # After clicking save, wait adaptively for dialog confirmation (fast then longer)
        try:
            dlg = _page_apo.wait_for_event("dialog", timeout=700)
            msg = dlg.message
            dlg.accept()
            if "Simpan Berhasil" in msg:
                return ("normal", msg)
            _page_apo.click(sel['reset_button'])
            return ("error", msg)
        except PWTimeoutError:
            # second-chance wait
            try:
                dlg = _page_apo.wait_for_event("dialog", timeout=3500)
                msg = dlg.message
                dlg.accept()
                if "Simpan Berhasil" in msg:
                    return ("normal", msg)
                _page_apo.click(sel['reset_button'])
                return ("error", msg)
            except PWTimeoutError:
                # final fallback: try to detect a known success element or assume no dialog -> error
                try:
                    # attempt to detect an element that normally contains success text if available
                    # if selector not present the evaluate returns false quickly
                    success_js = """(sel) => {
                        try {
                            const el = document.querySelector(sel);
                            if (!el) return false;
                            return el.textContent && el.textContent.includes('Simpan Berhasil');
                        } catch (e) {
                            return false;
                        }
                    }"""
                    # try quickly
                    success_detected = _adaptive_wait_for_function(
                        _page_apo,
                        success_js,
                        sel.get('success_text_selector', ''),  # optional selector in config
                        fast_timeout=300,
                        long_timeout=2000,
                        poll_interval=0.06
                    )
                    if success_detected:
                        return ("normal", "Simpan Berhasil (detected)")
                except Exception:
                    pass

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
