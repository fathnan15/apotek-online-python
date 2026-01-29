# apotek_runner.py
from playwright.sync_api import TimeoutError as PWTimeoutError, sync_playwright
from config import APOTEK_URL, APOTEK_SELECTORS
import time

_playwright_apo = None
_browser_apo    = None
_page_apo       = None
_last_sep       = None   # track last SEP to detect repeated SEP (cache issue)

def _now_ms():
    return int(time.time() * 1000)

def _adaptive_wait_for_function(page, js_func, arg, fast_timeout=800, long_timeout=7000, poll_interval=0.08):
    """
    Adaptive wait: fast path then longer path.
    Returns True if js_func(selector) returns truthy within time.
    """
    start = _now_ms()
    deadline_fast = start + fast_timeout
    deadline_long = start + long_timeout

    while _now_ms() <= deadline_fast:
        try:
            if page.evaluate(js_func, arg):
                return True
        except Exception:
            pass
        time.sleep(poll_interval)

    while _now_ms() <= deadline_long:
        try:
            if page.evaluate(js_func, arg):
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
    # keep default timeout moderate; adaptive waits handle longer cases
    _page_apo.set_default_timeout(5000)
    _page_apo.goto(APOTEK_URL, timeout=15000)
    print("✅ Connected to Apotek form.")

def submit_to_apotek(sep: str, receipt: str, rec_type: str) -> tuple[str, str]:
    """
    Perform: reset -> (maybe reload if same SEP) -> fill SEP -> wait for card -> fill receipt_type & receipt -> save -> handle dialog
    Returns (status, note)
    """
    global _last_sep, _page_apo
    sel = APOTEK_SELECTORS
    try:
        sep_str      = str(sep)
        receipt_str  = str(receipt)
        rec_type_str = str(rec_type)

        # Best-effort: clear the form before starting (help with stale/cached values)
        try:
            _page_apo.click(sel['reset_button'])
        except Exception:
            # not fatal; ignore
            pass
        _page_apo.wait_for_timeout(150)

        # If same SEP as last processed, do a page reload to avoid server-side caching returning previous patient data
        if _last_sep is not None and sep_str == _last_sep:
            try:
                _page_apo.reload(timeout=10000)
                # small pause after reload so scripts settle
                _page_apo.wait_for_timeout(300)
            except Exception:
                # reload failed — continue but results may be cached
                pass

        # Fill SEP and trigger search
        _page_apo.fill(sel['sep_input'], sep_str)
        _page_apo.keyboard.press("Enter")

        # Wait adaptively for the No Kartu input to have a non-empty value
        js_check_value = """(selector) => {
            try {
                const el = document.querySelector(selector);
                if (!el) return false;
                const v = el.value;
                return v !== null && v !== undefined && v.toString().trim().length > 0;
            } catch (e) {
                return false;
            }
        }"""
        # fast path ~0.7s, long path up to 7s — tuned for reliability
        ok = _adaptive_wait_for_function(
            _page_apo,
            js_check_value,
            sel['no_kartu_input'],
            fast_timeout=700,
            long_timeout=7000,
            poll_interval=0.06
        )

        if not ok:
            # try one forced reload as a last attempt (sometimes server returns stale DOM)
            try:
                _page_apo.reload(timeout=8000)
                _page_apo.wait_for_timeout(200)
                ok = _adaptive_wait_for_function(
                    _page_apo,
                    js_check_value,
                    sel['no_kartu_input'],
                    fast_timeout=500,
                    long_timeout=4000,
                    poll_interval=0.06
                )
            except Exception:
                ok = False

        if not ok:
            return ("error", "No card number returned by page")

        # Detect immediate dialog (error) - try a short then slightly longer wait
        try:
            dlg = _page_apo.wait_for_event("dialog", timeout=800)
            err_msg = dlg.message
            dlg.accept()
            try:
                _page_apo.click(sel['reset_button'])
            except Exception:
                pass
            return ("error", err_msg)
        except PWTimeoutError:
            try:
                dlg = _page_apo.wait_for_event("dialog", timeout=3000)
                err_msg = dlg.message
                dlg.accept()
                try:
                    _page_apo.click(sel['reset_button'])
                except Exception:
                    pass
                return ("error", err_msg)
            except PWTimeoutError:
                pass

        # Fill receipt type and receipt number
        _page_apo.fill(sel['receipt_type_input'], rec_type_str)
        _page_apo.fill(sel['receipt_input'], receipt_str)
        _page_apo.wait_for_timeout(150)
        _page_apo.click(sel['simpan_button'])

        # Wait for confirmation dialog (fast then fallback)
        try:
            dlg = _page_apo.wait_for_event("dialog", timeout=800)
            msg = dlg.message
            dlg.accept()
            if "Simpan Berhasil" in msg:
                _last_sep = sep_str
                return ("normal", msg)
            try:
                _page_apo.click(sel['reset_button'])
            except Exception:
                pass
            _last_sep = sep_str
            return ("error", msg)
        except PWTimeoutError:
            try:
                dlg = _page_apo.wait_for_event("dialog", timeout=4000)
                msg = dlg.message
                dlg.accept()
                if "Simpan Berhasil" in msg:
                    _last_sep = sep_str
                    return ("normal", msg)
                try:
                    _page_apo.click(sel['reset_button'])
                except Exception:
                    pass
                _last_sep = sep_str
                return ("error", msg)
            except PWTimeoutError:
                # final fallback: attempt to detect success text (optional selector)
                try:
                    success_js = """(sel) => {
                        try {
                            if (!sel) return false;
                            const el = document.querySelector(sel);
                            if (!el) return false;
                            return el.textContent && el.textContent.includes('Simpan Berhasil');
                        } catch (e) {
                            return false;
                        }
                    }"""
                    success_detected = False
                    sel_success = sel.get('success_text_selector', '')
                    if sel_success:
                        success_detected = _adaptive_wait_for_function(_page_apo, success_js, sel_success, fast_timeout=300, long_timeout=2000, poll_interval=0.06)
                    if success_detected:
                        _last_sep = sep_str
                        return ("normal", "Simpan Berhasil (detected)")
                except Exception:
                    pass

                try:
                    _page_apo.click(sel['reset_button'])
                except Exception:
                    pass
                _last_sep = sep_str
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
