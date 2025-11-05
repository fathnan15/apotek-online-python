# auto_input_obat_safe_patched.py
from playwright.sync_api import sync_playwright
import gspread
from google.oauth2.service_account import Credentials
from config import SERVICE_ACCOUNT_PATH
import time

# ==== RATE-LIMIT SAFE GOOGLE UPDATE HELPERS ====
import random
from gspread.exceptions import APIError
from datetime import datetime
import concurrent.futures

def safe_update_cell(ws, cell, value, retries=3):
    """
    Safe Google Sheets updater with retry & backoff to prevent 429 rate limit.
    Also writes a timestamp to column F when updating the resep sheet.
    """

    for attempt in range(retries):
        try:
            # primary update
            ws.update_acell(cell, value)

            # if this is the resep sheet, also write timestamp to column F of same row
            try:
                # match row number from cell (e.g. "G12" -> "12")
                row_digits = "".join(ch for ch in str(cell) if ch.isdigit())
                if row_digits and getattr(ws, "title", "").lower() == SHEET_RESEP.lower():
                    ts_cell = f"F{row_digits}"
                    ts_val = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    ws.update_acell(ts_cell, ts_val)
            except APIError as e_inner:
                # surface quota errors to outer handler so retry/backoff can happen
                if "Quota exceeded" in str(e_inner):
                    raise e_inner
                # otherwise don't block the main update; log and continue
                print(f"⚠️ Failed to write timestamp {ts_cell}: {e_inner}")

            time.sleep(1.1)  # Throttle slightly to stay under quota
            return True
        except APIError as e:
            if "Quota exceeded" in str(e):
                wait = 10 * (attempt + 1) + random.random() * 3
                print(f"⚠️ Quota exceeded. Cooling down {wait:.1f}s before retry...")
                time.sleep(wait)
            else:
                raise
    print(f"❌ Failed to update cell {cell} after {retries} retries.")
    return False

# --- NEW: small, deterministic writer wrapper that preserves ordering ---
def write_row_sync(ws, row, msg_text, kode_val):
    """
    Blocking wrapper that writes status and message to Google Sheet using safe_update_cell.
    Designed to be executed in a worker thread but the caller will wait for completion.
    Returns status string: "done" or "error".
    """
    msg = (msg_text or "").strip()
    msg_lower = msg.lower()
    try:
        # classify result
        if "obat berhasil disimpan" in msg_lower or "berhasil" in msg_lower:
            # success
            status = "done"
            print(f" ✅ 200-Success : Updating row {row} for obat {kode_val}")
        else:
            # non-success / message may contain error
            status = "error"
            print(f" ⚠️  Non-success : Updating row {row} for obat {kode_val} -> '{msg}'")

        # write status and message (H = status, I = message) - keep your previous layout
        safe_update_cell(ws, f"H{row}", status)
        safe_update_cell(ws, f"I{row}", msg)
        # return status for caller to inspect
        return status
    except Exception as e:
        print(f"❌ Exception while writing row {row}: {e}")
        return "error"

# ==== CONFIGURATION ====
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1MdEQrxNS6kuHkwks8Fgg6q29HxJ3qx2br-DPBpGecn4/edit?gid=1523826715#gid=1523826715"
SHEET_RESEP = "daftar resep"
SHEET_OBAT = "daftar obat"

CDP_ENDPOINT = "http://127.0.0.1:9222"
BASE_URL = "https://apotek.bpjs-kesehatan.go.id/apotek/"
SELECTORS = {
    "resep_filter": "#ctl00_ctl00_ASPxSplitter1_Content_ContentSplitter_MainContent_GvDaftarResep_DXFREditorcol13_I",
    "btn_input_obat": "#ctl00_ctl00_ASPxSplitter1_Content_ContentSplitter_MainContent_GvDaftarResep_cell0_11_BtnInputObat_CD",
    "kode_obat": "#ctl00_ctl00_ASPxSplitter1_Content_ContentSplitter_MainContent_TabPageObat_CboKdObatNR_I",
    "harga_obat": "#ctl00_ctl00_ASPxSplitter1_Content_ContentSplitter_MainContent_TabPageObat_TxtHrgTagObatNR_I",
    "qty_obat": "#ctl00_ctl00_ASPxSplitter1_Content_ContentSplitter_MainContent_TabPageObat_TxtJmlObatNR_I",
    "btn_simpan": "#ctl00_ctl00_ASPxSplitter1_Content_ContentSplitter_MainContent_TabPageObat_BtnSimpanNR_CD"
}

# ==== GOOGLE SHEET HANDLER ====
def open_sheet():
    SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_PATH, scopes=SCOPES)
    client = gspread.authorize(creds)
    ss = client.open_by_url(SPREADSHEET_URL)
    return ss.worksheet(SHEET_RESEP), ss.worksheet(SHEET_OBAT)

# ==== PLAYWRIGHT HELPERS ====
def attach_browser():
    pw = sync_playwright().start()
    browser = pw.chromium.connect_over_cdp(CDP_ENDPOINT)
    context = browser.contexts[0] if browser.contexts else browser.new_context()
    page = context.pages[0] if context.pages else context.new_page()
    print("✅ Attached to existing Chrome session.")
    return browser, page

def handle_dialog(page):
    try:
        dialog = page.wait_for_event("dialog", timeout=8000)
        msg = dialog.message
        dialog.accept()
        return msg
    except Exception:
        return None

def build_obat_row_map(ws_obat):
    """Build a mapping (receipt_num, apol_id) → row number for quick lookup."""
    values = ws_obat.get_all_values()
    if not values:
        return {}

    headers = [h.strip().lower() for h in values[0]]
    try:
        receipt_idx = headers.index("receipt_num")
        apol_idx = headers.index("apol_id")
        status_idx = headers.index("status")
    except ValueError:
        print(f"⚠️ Header mismatch. Headers found: {headers}")
        return {}

    mapping = {}
    for row_num, row in enumerate(values[1:], start=2):
        if len(row) <= max(receipt_idx, apol_idx):
            continue

        # Normalize keys: strip(), lowercase, and remove leading zeros
        no_resep = str(row[receipt_idx]).strip().replace("'", "").lstrip("0").lower()
        kode_obat = str(row[apol_idx]).strip().replace("'", "").lstrip("0").lower()
        status = str(row[status_idx]).strip().lower() if len(row) > status_idx else ""

        if no_resep and kode_obat and status not in ("normal","done", "error", "not_found", "checked","null"):
            mapping[(no_resep, kode_obat)] = row_num

    print(f"📊 Loaded {len(mapping)} obat rows into cache.")
    return mapping

# ==== MAIN ====
def auto_input():
    ws_resep, ws_obat = open_sheet()
    resep_records = ws_resep.get_all_records()
    obat_records = ws_obat.get_all_records()
    obat_row_map = build_obat_row_map(ws_obat)
    browser, page = attach_browser()

    # ThreadPoolExecutor reused for ordered background writes (we wait on each)
    with concurrent.futures.ThreadPoolExecutor(max_workers=2) as executor:
        for i, resep in enumerate(resep_records, start=2):
            status = str(resep.get("status", "")).strip().lower()
            if status in ("normal","done", "error", "not_found", "checked","null"):
                continue

            no_resep = str(resep.get("receipt_num", "")).strip()
            no_sep = str(resep.get("sep_num", "")).strip()
            if not no_resep:
                print(f"⚠️ Row {i} missing resep number.")
                continue

            print(f"\n🔎 Processing resep {no_resep} (SEP={no_sep})")

            related_obats = [
                o for o in obat_records
                if str(o.get("receipt_num", "")).strip() == no_resep
                and str(o.get("status", "")).strip().lower() not in ("normal","done", "error", "not_found", "checked","null")
            ]
            print(f"  📝 Found {len(related_obats)} pending obat for this resep.")
            if not related_obats:
                safe_update_cell(ws_resep, f"G{i}", "null")
                print(f"✅ Resep {no_resep} marked done (no pending obat).")
                continue

            page.goto(BASE_URL + "DaftarResep.aspx")
            page.wait_for_load_state("networkidle")
            page.fill(SELECTORS["resep_filter"], no_resep)
            page.keyboard.press("Enter")

            try:
                page.wait_for_selector(f"text={no_resep}", timeout=15000)
            except Exception:
                print(f"❌ Resep {no_resep} not found in table.")
                safe_update_cell(ws_resep, f"G{i}", "not_found")
                continue

            print("🕐 Clicking Input Obat button…")
            # Wait until grid finishes loading before clicking
            try:
                # Wait for overlay to appear and then disappear
                page.wait_for_selector("div.dxgvLoadingDiv_Glass", state="visible", timeout=5000)
                page.wait_for_selector("div.dxgvLoadingDiv_Glass", state="hidden", timeout=15000)
            except:
                # Overlay might not appear at all (already loaded)
                pass

            # Re-locate the button (old handles may be detached)
            buttons = page.query_selector_all(SELECTORS["btn_input_obat"])
            if not buttons:
                print(f"❌ No Input Obat button found for resep {no_resep}")
                continue

            # Now click safely
            buttons[0].click()

            # ⏳ Wait until redirected to ObatInput.aspx (instead of fixed sleep)
            try:
                page.wait_for_url("**/ObatInput.aspx", timeout=30000)
                page.wait_for_load_state("networkidle")
                print("✅ ObatInput.aspx fully loaded.")
            except Exception:
                print("⚠️ Timeout waiting for ObatInput.aspx, continue anyway.")

            # Track if any obat for this resep produced an error
            resep_has_error = False

            for obat in related_obats:
                kode = str(obat.get("apol_id", "")).strip()
                qty = str(obat.get("qty", "")).strip() or "1"
                if not kode:
                    continue

                print(f"  💊 Inputting {kode} x{qty} …")

                # --- robust autocomplete selection (replacement) ---
                LISTBOX_SELECTOR = "table[id$='CboKdObatNR_DDD_L_LBT']"
                FIRST_ITEM_ROW = "table[id$='CboKdObatNR_DDD_L_LBT'] tr.dxeListBoxItemRow_Glass"
                FIRST_ITEM_KD_CELL = "table[id$='CboKdObatNR_DDD_L_LBT'] td[id$='_LBI0T0']"

                def read_kode_input_value():
                    try:
                        return page.eval_on_selector(SELECTORS["kode_obat"], "el => el.value").strip()
                    except Exception:
                        return ""

                # type to trigger autocomplete
                page.fill(SELECTORS["kode_obat"], "")
                time.sleep(0.12)
                page.click(SELECTORS["kode_obat"])
                page.type(SELECTORS["kode_obat"], kode, delay=50)

                # wait for the listbox to appear and try to click first item
                try:
                    page.wait_for_selector(LISTBOX_SELECTOR, timeout=6000)
                    try:
                        page.click(FIRST_ITEM_KD_CELL, timeout=3000)
                    except Exception:
                        try:
                            page.click(FIRST_ITEM_ROW, timeout=3000)
                        except Exception:
                            page.keyboard.press("ArrowDown")
                            time.sleep(0.18)
                            page.keyboard.press("Enter")
                except Exception:
                    # listbox never showed — fallback to ArrowDown/Enter
                    time.sleep(0.9)
                    page.keyboard.press("ArrowDown")
                    time.sleep(0.18)
                    page.keyboard.press("Enter")

                # short pause to let widget propagate selection to fields
                time.sleep(0.5)

                # verify selection
                selected_val = read_kode_input_value()
                if selected_val and (kode in selected_val or selected_val in kode):
                    ui_ok = True
                else:
                    try:
                        found = page.query_selector(f"xpath=//table[contains(@id,'TabPageObat')]//td[contains(., '{kode}')]")
                        ui_ok = bool(found)
                    except:
                        ui_ok = False

                if not ui_ok:
                    print(f"⚠️ Autocomplete selection for {kode} may have failed — selected_val='{selected_val}'. Will attempt one retry.")
                    # single retry
                    page.fill(SELECTORS["kode_obat"], "")
                    time.sleep(0.12)
                    page.click(SELECTORS["kode_obat"])
                    page.type(SELECTORS["kode_obat"], kode, delay=80)
                    try:
                        page.wait_for_selector(LISTBOX_SELECTOR, timeout=5000)
                        page.click(FIRST_ITEM_KD_CELL)
                    except:
                        page.keyboard.press("ArrowDown")
                        page.keyboard.press("Enter")
                    time.sleep(0.6)
                    selected_val = read_kode_input_value()
                    ui_ok = (selected_val and (kode in selected_val or selected_val in kode))

                if not ui_ok:
                    print(f"❌ Failed to reliably select kode {kode}. Selected value after retry: '{selected_val}'. Skipping this obat for now.")
                    # Mark as error in sheet optionally (we skip for now)
                    resep_has_error = True
                    continue

                # proceed to fill qty & save as before
                time.sleep(0.2)
                page.fill(SELECTORS["qty_obat"], qty)
                page.click(SELECTORS["btn_simpan"])

                message = handle_dialog(page)
                print(f"💬 {message or 'No alert dialog detected.'}")

                # Update Google Sheet immediately (run in thread but wait here to preserve ordering)
                row = obat_row_map.get((no_resep, kode))

                if not row:
                    print(f"DEBUG: lookup key=({no_resep}, {kode})")
                    print("DEBUG: available keys (sample):", list(obat_row_map.keys())[:5])

                if row:
                    # Submit to thread executor and wait for completion before moving on
                    future = executor.submit(write_row_sync, ws_obat, row, message or "", kode)
                    try:
                        status_result = future.result(timeout=120)  # wait for write to finish
                        if status_result == "done":
                            print(f"  ✅ Completed write for row {row} (obat {kode})")
                        else:
                            print(f"  ⚠️ Write returned status '{status_result}' for row {row} (obat {kode})")
                            resep_has_error = True
                    except concurrent.futures.TimeoutError:
                        print(f"❌ Timeout while writing row {row} to sheet.")
                        resep_has_error = True
                else:
                    print(f"⚠️ Could not find row for resep {no_resep}, obat {kode}")
                    resep_has_error = True

                time.sleep(1)

            # After processing all obat for this resep, set resep status depending on any obat errors
            final_status = "error" if resep_has_error else "done"
            safe_update_cell(ws_resep, f"G{i}", final_status)
            print(f"✅ Resep {no_resep} completed. Final status: {final_status.upper()}")
            time.sleep(2.5)

    browser.close()
    print("🏁 All resep processed safely and completely.")

if __name__ == "__main__":
    if input("Enter sheet name for resep (or leave blank for default 'daftar resep'): ").strip():
        SHEET_RESEP = input("Sheet Name for Resep (e.g. daftar resep): ").strip()
    auto_input()
