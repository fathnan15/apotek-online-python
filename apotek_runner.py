from playwright.sync_api import TimeoutError as PWTimeoutError, sync_playwright
from config import APOTEK_URL, APOTEK_SELECTORS
import time

_playwright_apo = None
_browser_apo    = None
_page_apo       = None
_last_sep       = None 

def _now_ms():
    return int(time.time() * 1000)

def _adaptive_wait_for_function(page, js_func, arg, fast_timeout=800, long_timeout=7000, poll_interval=0.08):
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

def init_apotek(cdp_endpoint: str = "http://127.0.0.1:9515"):
    global _playwright_apo, _browser_apo, _page_apo
    _playwright_apo = sync_playwright().start()
    _browser_apo    = _playwright_apo.chromium.connect_over_cdp(cdp_endpoint)
    ctx             = _browser_apo.contexts[0] if _browser_apo.contexts else _browser_apo.new_context()
    _page_apo       = ctx.pages[0] if ctx.pages else ctx.new_page()
    _page_apo.set_default_timeout(5000)
    
    try:
        if "apotek" not in _page_apo.url:
            _page_apo.goto(APOTEK_URL, timeout=15000)
    except:
        _page_apo.goto(APOTEK_URL, timeout=15000)
        
    print("✅ Connected to Apotek form.")

def close_apotek():
    global _browser_apo, _playwright_apo
    if _browser_apo:
        _browser_apo.close()
    if _playwright_apo:
        _playwright_apo.stop()

# =========================================================
# HELPER: Pre-Save Verification
# =========================================================
def _verify_and_fix_fields(page, receipt_sel, expected_receipt, type_sel, expected_type):
    """
    Checks if the fields contain the expected values.
    If not, refills them.
    Repeats until correct or retries exhausted.
    """
    max_retries = 5
    for i in range(max_retries):
        # 1. Read current values
        try:
            curr_receipt = str(page.input_value(receipt_sel)).strip()
            # For dropdowns/combos, input_value usually grabs the text in the box
            curr_type    = str(page.input_value(type_sel)).strip()
        except Exception:
            curr_receipt = ""
            curr_type    = ""

        # 2. Compare
        receipt_match = (curr_receipt == expected_receipt)
        # We do a loose check for type because sometimes dropdowns have extra whitespace
        type_match    = (expected_type in curr_type) if curr_type else False

        if receipt_match and type_match:
            return True # All good, proceed to save

        # 3. Fix Mismatches
        print(f"⚠️ Mismatch detected (Attempt {i+1}/{max_retries}). Fixing...")
        
        if not type_match:
            try:
                page.fill(type_sel, expected_type)
                page.wait_for_timeout(100)
            except: pass

        if not receipt_match:
            try:
                page.click(receipt_sel) # Focus first
                page.fill(receipt_sel, expected_receipt)
            except: pass
        
        # Wait for value to settle before next check
        page.wait_for_timeout(300)

    return False # Failed to stabilize

# =========================================================
# MAIN LOGIC
# =========================================================
def submit_revision_to_apotek(sep: str, receipt: str, rec_type: str, date_str: str, iter_stat: str | None) -> tuple[str, str]:
    return _internal_submit(sep, receipt, rec_type, date_str, iter_stat)

def submit_to_apotek(sep: str, receipt: str, rec_type: str) -> tuple[str, str]:
    return _internal_submit(sep, receipt, rec_type, None, None)

def _internal_submit(sep: str, receipt: str, rec_type: str, date_str: str | None, iter_stat: str | None) -> tuple[str, str]:
    global _last_sep, _page_apo
    sel = APOTEK_SELECTORS
    try:
        sep_str      = str(sep).strip()
        receipt_str  = str(receipt).strip()
        rec_type_str = str(rec_type).strip()

        # 1. Reset
        try:
            _page_apo.click(sel['reset_button'])
        except Exception:
            pass
        _page_apo.wait_for_timeout(200)

        # 2. Fill SEP
        _page_apo.fill(sel['sep_input'], sep_str)
        _page_apo.keyboard.press("Enter")

        # 3. Wait for Patient Data (implies SEP valid)
        js_check_value = """(selector) => {
            try {
                const el = document.querySelector(selector);
                if (!el) return false;
                const v = el.value;
                return v !== null && v !== undefined && v.toString().trim().length > 0;
            } catch (e) { return false; }
        }"""
        
        ok = _adaptive_wait_for_function(_page_apo, js_check_value, sel['no_kartu_input'], fast_timeout=800, long_timeout=5000)
        
        if not ok:
            # Retry reload once
            try:
                _page_apo.reload(timeout=5000)
                _page_apo.wait_for_timeout(500)
                _page_apo.fill(sel['sep_input'], sep_str)
                _page_apo.keyboard.press("Enter")
                ok = _adaptive_wait_for_function(_page_apo, js_check_value, sel['no_kartu_input'], fast_timeout=800, long_timeout=5000)
            except: pass

        if not ok:
            # Check for error dialog
            try:
                dlg = _page_apo.wait_for_event("dialog", timeout=500)
                msg = dlg.message
                dlg.accept()
                try: _page_apo.click(sel['reset_button'])
                except: pass
                return ("error", msg)
            except PWTimeoutError:
                return ("error", "SEP load timeout (No card number)")

        # 4. Fill Dates (if revision)
        if date_str:
            _page_apo.wait_for_timeout(200)
            _page_apo.fill(sel['date_receipt_input'], date_str)
            _page_apo.keyboard.press("Tab")
            _page_apo.wait_for_timeout(200)
            
            _page_apo.fill(sel['date_service_input'], date_str)
            _page_apo.keyboard.press("Tab")
            # Wait for any AJAX caused by date change
            _page_apo.wait_for_timeout(400)


        # 5. Fill Receipt Info (Initial Attempt)
        _page_apo.fill(sel['receipt_type_input'], rec_type_str)
        _page_apo.fill(sel['receipt_input'], receipt_str)
        _page_apo.wait_for_timeout(200)

        # If iteration status element exists, check if it shows "Iterasi"
        if iter_stat is not None:
            try:
                # _page_apo.wait_for_timeout(200)
                # _page_apo.fill(sel['iteration_status'], "")
                _page_apo.wait_for_timeout(200)
                _page_apo.fill(sel['iteration_status'], iter_stat).strip()
                _page_apo.keyboard.press("Tab")
                _page_apo.wait_for_timeout(200)
            except Exception:
                pass
        # ======================================================
        # 6. CRITICAL: PRE-SAVE VERIFICATION
        # ======================================================
        # Check if values are actually there before clicking Simpan
        is_valid = _verify_and_fix_fields(
            _page_apo, 
            sel['receipt_input'], receipt_str,
            sel['receipt_type_input'], rec_type_str
        )
        
        if not is_valid:
             return ("error", "Verification failed: Fields keep reverting or blocked.")

        # ======================================================
        # 7 & 8. Submit and Wait for Result
        # ======================================================
        try:
            # Use expect_event BEFORE clicking to prevent race conditions
            # Extended timeout to 15s to account for slow server responses
            with _page_apo.expect_event("dialog", timeout=15000) as dialog_info:
                _page_apo.click(sel['simpan_button'])
            
            dlg = dialog_info.value
            msg = dlg.message
            dlg.accept()
            
            # Use a slightly broader success check
            if "Simpan Berhasil" in msg or "berhasil" in msg.lower():
                _last_sep = sep_str
                return ("normal", msg)
            
            try: _page_apo.click(sel['reset_button'])
            except: pass
            _last_sep = sep_str
            return ("error", msg)

        except PWTimeoutError:
            try: _page_apo.click(sel['reset_button'])
            except: pass
            return ("error", "No confirmation dialog (Timeout after 15s)")

    except Exception as e:
        return ("error", str(e))