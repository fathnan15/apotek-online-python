from playwright.sync_api import sync_playwright
import gspread
from google.oauth2.service_account import Credentials
from config import SERVICE_ACCOUNT_PATH
import time

# ==== RATE-LIMIT SAFE GOOGLE UPDATE HELPERS ====
import random
from gspread.exceptions import APIError

def safe_update_cell(ws, cell, value, retries=3):
    """
    Safe Google Sheets updater with retry & backoff to prevent 429 rate limit.
    """
    for attempt in range(retries):
        try:
            ws.update_acell(cell, value)
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
    context = browser.contexts[0]
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

        if no_resep and kode_obat and status not in ("done", "ok", "selesai"):
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

    for i, resep in enumerate(resep_records, start=2):
        status = str(resep.get("status", "")).strip().lower()
        if status in ("done", "ok", "selesai"):
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
            and str(o.get("status", "")).strip().lower() not in ("done", "ok", "selesai")
        ]
        print(f"  📝 Found {len(related_obats)} pending obat for this resep.")
        if not related_obats:
            # safe_update_cell(ws_resep, f"G{i}", "")
            print(f"✅ Resep {no_resep} marked done (no pending obat).")
            continue

        # NAVIGATE to DaftarResep and wait specifically for the filter input (more precise than networkidle)
        page.goto(BASE_URL + "DaftarResep.aspx")
        try:
            page.locator(SELECTORS["resep_filter"]).wait_for(timeout=10000)
        except Exception:
            # fallback: small wait if the filter didn't appear quickly
            time.sleep(1)

        # fill filter and press Enter; then wait for the resep text to appear
        page.fill(SELECTORS["resep_filter"], no_resep)
        page.keyboard.press("Enter")
        try:
            page.wait_for_selector(f"text={no_resep}", timeout=15000)
        except Exception:
            print(f"❌ Resep {no_resep} not found in table.")
            safe_update_cell(ws_resep, f"G{i}", "not_found")
            continue

        print("🕐 Waiting for Input Obat button to be ready…")
        # Wait for the Input Obat button locator to be attached & visible (avoid stale handles)
        try:
            btn_locator = page.locator(SELECTORS["btn_input_obat"]).first
            btn_locator.wait_for(state="visible", timeout=10000)
            # wait for any loading overlay to disappear (shorter timeouts)
            try:
                page.wait_for_selector("div.dxgvLoadingDiv_Glass", state="hidden", timeout=8000)
            except:
                pass
            btn_locator.click()
        except Exception as e:
            print(f"❌ Could not click Input Obat button for resep {no_resep}: {e}")
            continue

        # Wait for ObatInput.aspx by URL pattern (keeps you from using long sleeps)
        try:
            page.wait_for_url("**/ObatInput.aspx", timeout=25000)
            # ensure kode input is ready
            page.locator(SELECTORS["kode_obat"]).wait_for(state="visible", timeout=10000)
            print("✅ ObatInput.aspx loaded.")
        except Exception:
            print("⚠️ ObatInput.aspx might not have loaded fully (continuing nonetheless).")

        # Pre-define listbox selectors used in autocomplete flow
        LISTBOX_SELECTOR = "table[id$='CboKdObatNR_DDD_L_LBT']"
        FIRST_ITEM_ROW = "table[id$='CboKdObatNR_DDD_L_LBT'] tr.dxeListBoxItemRow_Glass"
        FIRST_ITEM_KD_CELL = "table[id$='CboKdObatNR_DDD_L_LBT'] td[id$='_LBI0T0']"

        def read_kode_input_value():
            try:
                return page.eval_on_selector(SELECTORS["kode_obat"], "el => el.value").strip()
            except Exception:
                return ""

        for obat in related_obats:
            kode = str(obat.get("apol_id", "")).strip()
            qty = str(obat.get("qty", "")).strip() or "1"
            if not kode:
                continue

            print(f"  💊 Inputting {kode} x{qty} …")

            # --- robust autocomplete selection (optimized timing) ---
            page.fill(SELECTORS["kode_obat"], "")
            # tiny pause to ensure clear has propagated
            time.sleep(0.08)

            # ensure the input is focused using locator.click (less brittle)
            kode_loc = page.locator(SELECTORS["kode_obat"])
            kode_loc.click()
            # faster typing but still human-like; reduced delay to speed up
            page.type(SELECTORS["kode_obat"], kode, delay=60)

            # WAIT: prefer direct click on list item if it appears — faster & reliable
            try:
                page.wait_for_selector(LISTBOX_SELECTOR, timeout=4500)
                # try clicking the first cell (KD cell) then the row
                try:
                    page.locator(FIRST_ITEM_KD_CELL).first.click(timeout=2000)
                except Exception:
                    try:
                        page.locator(FIRST_ITEM_ROW).first.click(timeout=2000)
                    except Exception:
                        # fallback to keyboard selection
                        page.keyboard.press("ArrowDown")
                        time.sleep(0.08)
                        page.keyboard.press("Enter")
            except Exception:
                # listbox didn't show — fallback to keyboard selection with brief wait
                time.sleep(0.45)
                page.keyboard.press("ArrowDown")
                time.sleep(0.08)
                page.keyboard.press("Enter")

            # short verification pause (reduced)
            time.sleep(0.35)

            selected_val = read_kode_input_value()
            ui_ok = bool(selected_val and (kode in selected_val or selected_val in kode))
            if not ui_ok:
                # secondary DOM presence check (cheap)
                try:
                    found = page.query_selector(f"xpath=//table[contains(@id,'TabPageObat')]//td[contains(., '{kode}')]")
                    ui_ok = bool(found)
                except:
                    ui_ok = False

            if not ui_ok:
                print(f"⚠️ Autocomplete selection for {kode} may have failed — selected_val='{selected_val}'. Retrying once.")
                # single retry (light)
                page.fill(SELECTORS["kode_obat"], "")
                time.sleep(0.08)
                kode_loc.click()
                page.type(SELECTORS["kode_obat"], kode, delay=70)
                try:
                    page.wait_for_selector(LISTBOX_SELECTOR, timeout=3500)
                    page.locator(FIRST_ITEM_KD_CELL).first.click(timeout=1500)
                except:
                    page.keyboard.press("ArrowDown")
                    page.keyboard.press("Enter")
                time.sleep(0.5)
                selected_val = read_kode_input_value()
                ui_ok = bool(selected_val and (kode in selected_val or selected_val in kode))

            if not ui_ok:
                print(f"❌ Failed to reliably select kode {kode}. Selected value after retry: '{selected_val}'. Skipping this obat for now.")
                continue

            # Fill qty & save (kept intact), with small throttle
            time.sleep(0.5)
            page.fill(SELECTORS["qty_obat"], qty)
            page.click(SELECTORS["btn_simpan"])
            message = handle_dialog(page)
            print(f"💬 {message or 'No alert dialog detected.'}")

            # Update Google Sheet immediately
            row = obat_row_map.get((no_resep, kode))
            if not row:
                print(f"DEBUG: lookup key=({no_resep}, {kode})")
                print("DEBUG: available keys (sample):", list(obat_row_map.keys())[:5])

            if row:
                safe_update_cell(ws_obat, f"H{row}", message or "done")
                print(f"  ✅ Updated row {row} for obat {kode}")
            else:
                print(f"⚠️ Could not find row for resep {no_resep}, obat {kode}")

            # short cooldown between obat items (reduced)
            time.sleep(0.9)

        # mark resep done (using safe writer)
        safe_update_cell(ws_resep, f"G{i}", "done")
        print(f"✅ Resep {no_resep} completed.")
        time.sleep(1.5)  # shorter between resep

    browser.close()
    print("🏁 All resep processed safely and completely.")

if __name__ == "__main__":
    auto_input()
