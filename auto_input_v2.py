# auto_input_obat_v9_full_revise.py
import argparse
import threading
import time
import random
from datetime import datetime
import concurrent.futures

from playwright.sync_api import sync_playwright
import gspread
from google.oauth2.service_account import Credentials
from gspread.exceptions import APIError

from config import SERVICE_ACCOUNT_PATH

# ==== CONFIGURATION ====
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1MdEQrxNS6kuHkwks8Fgg6q29HxJ3qx2br-DPBpGecn4/edit?gid=1523826715#gid=1523826715"
SHEET_RESEP = "daftar resep"
SHEET_OBAT = "daftar obat"

BASE_URL = "https://apotek.bpjs-kesehatan.go.id/apotek/"
SELECTORS = {
    "resep_filter": "#ctl00_ctl00_ASPxSplitter1_Content_ContentSplitter_MainContent_GvDaftarResep_DXFREditorcol13_I",
    "sep_filter": "#ctl00_ctl00_ASPxSplitter1_Content_ContentSplitter_MainContent_GvDaftarResep_DXFREditorcol3_I",
    "btn_input_obat": "#ctl00_ctl00_ASPxSplitter1_Content_ContentSplitter_MainContent_GvDaftarResep_cell0_11_BtnInputObat_CD",
    "kode_obat": "#ctl00_ctl00_ASPxSplitter1_Content_ContentSplitter_MainContent_TabPageObat_CboKdObatNR_I",
    "qty_obat": "#ctl00_ctl00_ASPxSplitter1_Content_ContentSplitter_MainContent_TabPageObat_TxtJmlObatNR_I",
    "btn_simpan": "#ctl00_ctl00_ASPxSplitter1_Content_ContentSplitter_MainContent_TabPageObat_BtnSimpanNR_CD",
    "table_obat": "#ctl00_ctl00_ASPxSplitter1_Content_ContentSplitter_MainContent_GvObat_DXMainTable",
    "loading_overlay": ".dxgvLoadingDiv_Glass"
}

# ==== ARGUMENT PARSING ====
def parse_args():
    parser = argparse.ArgumentParser(description="BPJS Auto Input Worker")
    parser.add_argument("--port", type=str, required=True, help="Chrome CDP Port (e.g., 9222)")
    parser.add_argument("--start-row", type=int, required=True, help="First row to process")
    parser.add_argument("--end-row", type=int, required=True, help="Last row to process")
    return parser.parse_args()

# ==== BATCH UPDATER SYSTEM ====
class SheetBatchUpdater:
    def __init__(self, ws, batch_size=20):
        self.ws = ws
        self.batch_size = batch_size
        self.updates = []
        self.lock = threading.Lock()

    def add_update(self, range_name, values):
        with self.lock:
            self.updates.append({'range': range_name, 'values': [values]})
            if len(self.updates) >= self.batch_size:
                self._execute_flush()

    def flush(self):
        with self.lock:
            self._execute_flush()

    def _execute_flush(self):
        if not self.updates:
            return
            
        for attempt in range(4):
            try:
                self.ws.batch_update(self.updates)
                print(f"📦 Flushed {len(self.updates)} updates to '{self.ws.title}' successfully.")
                self.updates = []
                return
            except APIError as e:
                if "Quota exceeded" in str(e):
                    wait = (attempt + 1) * 5 + random.random() * 2
                    print(f"⚠️ Quota exceeded during batch flush. Cooling down {wait:.1f}s...")
                    time.sleep(wait)
                else:
                    print(f"❌ Unrecoverable API Error during flush: {e}")
                    return
            except Exception as e:
                print(f"❌ Network/Unknown error during flush: {e}")
                time.sleep(2)
        print("❌ Failed to flush batch after 4 attempts. Data may be out of sync.")

# ==== GOOGLE SHEET HANDLER ====
def open_sheet():
    SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_PATH, scopes=SCOPES)
    client = gspread.authorize(creds)
    ss = client.open_by_url(SPREADSHEET_URL)
    return ss.worksheet(SHEET_RESEP), ss.worksheet(SHEET_OBAT)

# ==== UTILS ====
def attach_browser(port):
    pw = sync_playwright().start()
    cdp_url = f"http://127.0.0.1:{port}"
    browser = pw.chromium.connect_over_cdp(cdp_url)
    context = browser.contexts[0] if browser.contexts else browser.new_context()
    page = context.pages[0] if context.pages else context.new_page()
    return browser, page, pw

def handle_dialog(page):
    try:
        dialog = page.wait_for_event("dialog", timeout=5000)
        msg = dialog.message
        dialog.accept()
        return msg
    except:
        return None

def build_obat_row_map(ws_obat):
    values = ws_obat.get_all_values()
    if not values: return {}
    
    headers = [h.strip().lower() for h in values[0]]
    try:
        r_idx = headers.index("receipt_num")
        p_idx = headers.index("presc_id")
        a_idx = headers.index("apol_id")
    except ValueError as e:
        print(f"⚠️ Header missing: {e}. Check sheet headers.")
        return {}

    mapping = {}
    for row_num, row in enumerate(values[1:], start=2):
        if len(row) <= max(r_idx, p_idx, a_idx): continue
        r_val = str(row[r_idx]).strip().replace("'", "").zfill(5).lower()
        p_val = str(row[p_idx]).strip().lstrip("0").lower()
        a_val = str(row[a_idx]).strip().replace("'", "").lstrip("0").lower()
        
        if r_val and a_val:
            mapping[(r_val, p_val, a_val)] = row_num
            
    return mapping

def wait_for_loading_gone(page):
    try:
        page.wait_for_selector(SELECTORS["loading_overlay"], state="hidden", timeout=10000)
        time.sleep(0.5)
    except Exception:
        pass

def queue_obat_update(updater, row, msg_text, kode_val):
    msg = (msg_text or "").strip()
    msg_lower = msg.lower()
    
    if any(x in msg_lower for x in ["berhasil", "already input"]):
        status = "done"
        print(f" ✅ Success: Row {row} (Obat {kode_val}) -> {status}")
    else:
        status = "error"
        print(f" ⚠️  Error/Message: Row {row} (Obat {kode_val}) -> '{msg}'")

    updater.add_update(f"H{row}:I{row}", [status, msg])
    return status

# ==== MAIN LOGIC ====
def auto_input(args):
    ws_resep, ws_obat = open_sheet()
    
    # Initialize batch updaters
    resep_updater = SheetBatchUpdater(ws_resep, batch_size=20)
    obat_updater = SheetBatchUpdater(ws_obat, batch_size=20)
    
    resep_records = ws_resep.get_all_records()
    obat_records = ws_obat.get_all_records()
    obat_row_map = build_obat_row_map(ws_obat)
    
    browser, page, pw = attach_browser(args.port)

    try:
        with concurrent.futures.ThreadPoolExecutor(max_workers=2) as executor:
            for i, resep in enumerate(resep_records, start=2):
                
                # Partitioning Enforcement
                if i < args.start_row or i > args.end_row:
                    continue

                status = str(resep.get("status", "")).strip().lower()
                if status in ("done", "error", "not_found", "null"):
                    continue
                    
                is_revise = (status == "revise")

                no_resep = str(resep.get("receipt_num", "")).strip().zfill(5)
                no_sep = str(resep.get("sep_num", "")).strip()
                presc_id = str(resep.get("presc_id", "")).strip().lstrip("0")

                if not no_resep: continue

                print(f"\n🔎 Processing Resep: {no_resep} (SEP: {no_sep}, Presc: {presc_id})")

                # REVISION FIX: If revise, exclude nothing (). If normal, exclude done/error/null.
                excluded_obat_statuses = () if is_revise else ("done", "error", "null")

                related_obats = [
                    o for o in obat_records
                    if str(o.get("receipt_num", "")).strip().zfill(5) == no_resep
                    and str(o.get("presc_id", "")).strip().lstrip("0") == presc_id
                    and str(o.get("status", "")).strip().lower() not in excluded_obat_statuses
                ]
                
                if not related_obats:
                    print("  ✅ No pending drugs for this resep.")
                    resep_updater.add_update(f"G{i}:H{i}", ["", "done"])
                    continue

                # ==========================================
                # 1. NAVIGATION PHASE
                # ==========================================
                nav_success = False
                try:
                    page.goto(BASE_URL + "DaftarResep.aspx")
                    page.wait_for_load_state("networkidle")
                    
                    if no_sep:
                        page.fill(SELECTORS["sep_filter"], no_sep)
                        page.keyboard.press("Enter")
                        wait_for_loading_gone(page)
                    
                    page.fill(SELECTORS["resep_filter"], no_resep)
                    page.keyboard.press("Enter")
                    wait_for_loading_gone(page)
                    
                    try:
                        page.wait_for_selector(f"text={no_resep}", timeout=10000)
                    except:
                        print(f"❌ Resep {no_resep} not found in grid.")
                        resep_updater.add_update(f"G{i}:H{i}", ["", "not_found"])
                        continue

                    wait_for_loading_gone(page)
                    
                    for attempt in range(1, 4):
                        btn = page.query_selector(SELECTORS["btn_input_obat"])
                        if not btn:
                            time.sleep(2)
                            continue
                        
                        try:
                            btn.click(force=True)
                            page.wait_for_url("**/ObatInput.aspx", timeout=6000)
                            page.wait_for_load_state("networkidle")
                            nav_success = True
                            break
                        except Exception:
                            wait_for_loading_gone(page)
                            time.sleep(1)

                    if not nav_success:
                        print(f"❌ Failed to enter input page after 3 attempts.")
                        resep_updater.add_update(f"G{i}:H{i}", ["", "error"])
                        continue

                except Exception as e:
                    print(f"  ❌ Navigation Error: {e}")
                    resep_updater.add_update(f"G{i}:H{i}", ["", "error"])
                    continue

                # ==========================================
                # 2. INPUT PHASE
                # ==========================================
                existing_map = {}
                try:
                    page.wait_for_selector(SELECTORS["table_obat"], state="visible", timeout=5000)
                    rows = page.query_selector_all(f"{SELECTORS['table_obat']} tr.dxgvDataRow_Glass")
                    for r in rows:
                        cells = r.query_selector_all("td")
                        if len(cells) > 7:
                            k_txt = cells[1].inner_text().strip().lstrip("0")
                            q_txt = cells[7].inner_text().strip()
                            if k_txt:
                                existing_map[k_txt] = q_txt
                except Exception:
                    pass

                resep_has_error = False
                for obat in related_obats:
                    kode_raw = str(obat.get("apol_id", "")).strip()
                    kode_clean = kode_raw.lstrip("0")
                    qty = str(obat.get("qty", "")).strip()
                    
                    row_key = (no_resep.lower(), presc_id.lower(), kode_clean.lower())
                    row_num = obat_row_map.get(row_key)
                    
                    if not row_num:
                        print(f"  ⚠️ Could not find sheet row for {kode_raw}")
                        continue

                    if kode_clean in existing_map:
                        existing_qty = existing_map[kode_clean]
                        msg = f"already input qty : {existing_qty}"
                        executor.submit(queue_obat_update, obat_updater, row_num, msg, kode_raw)
                        continue

                    print(f"  💊 Inputting {kode_raw} x{qty}...")
                    try:
                        page.fill(SELECTORS["kode_obat"], "")
                        time.sleep(0.1)
                        page.type(SELECTORS["kode_obat"], kode_raw, delay=50)
                        time.sleep(0.5)
                        
                        try:
                            page.keyboard.press("ArrowDown")
                            page.keyboard.press("Enter")
                        except: pass
                        
                        qty_ok = False
                        for q_attempt in range(3):
                            page.fill(SELECTORS["qty_obat"], "")
                            time.sleep(0.1)
                            page.type(SELECTORS["qty_obat"], qty, delay=50)
                            time.sleep(0.2)
                            
                            filled_val = page.eval_on_selector(SELECTORS["qty_obat"], "el => el.value").strip()
                            
                            if filled_val == qty:
                                qty_ok = True
                                break
                            else:
                                time.sleep(0.5)
                        
                        if not qty_ok:
                            err_msg = f"Qty Mismatch Failed: {filled_val} != {qty}"
                            queue_obat_update(obat_updater, row_num, err_msg, kode_raw)
                            resep_has_error = True
                            continue

                        page.click(SELECTORS["btn_simpan"])
                        msg_alert = handle_dialog(page) or ""
                        
                        future = executor.submit(queue_obat_update, obat_updater, row_num, msg_alert, kode_raw)
                        if future.result() == "error":
                            resep_has_error = True
                        time.sleep(1) 

                    except Exception as e:
                        print(f"  ❌ Error inputting {kode_raw}: {e}")
                        queue_obat_update(obat_updater, row_num, f"Script Error: {str(e)}", kode_raw)
                        resep_has_error = True

                final_status = "error" if resep_has_error else "done"
                ts_val = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                resep_updater.add_update(f"G{i}:H{i}", [ts_val, final_status])
                print(f"✅ Resep {no_resep} finished -> {final_status.upper()}")

    finally:
        print("\n🛑 Script terminating. Executing final data flush...")
        resep_updater.flush()
        obat_updater.flush()
        browser.close()
        pw.stop()

if __name__ == "__main__":
    args = parse_args()
    auto_input(args)