import gspread
import socket
import time
import uuid
from datetime import datetime, timedelta
from google.oauth2.service_account import Credentials
from config import (
    SERVICE_ACCOUNT_PATH,
    SHEET_URL,
    SEP_SHEET_HEADERS
)

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
HOSTNAME = socket.gethostname()

def get_worksheet(name: str):
    """Authenticate and open the named worksheet."""
    creds  = Credentials.from_service_account_file(SERVICE_ACCOUNT_PATH, scopes=SCOPES)
    client = gspread.authorize(creds)
    sheet  = client.open_by_url(SHEET_URL)
    return sheet.worksheet(name)

def read_all_records(ws) -> list[dict]:
    """Reads all records from the sheet into a list of dicts."""
    values = ws.get_all_values()
    if not values:
        return []
    headers = values[0]
    data    = values[1:]
    # Basic safety: ensure we don't crash if empty rows exist
    records = [dict(zip(headers, row)) for row in data if row]
    return records

def _now_iso():
    return datetime.now().strftime("%Y-%m-%dT%H:%M:%S")

# ==========================================
# 1. EXTRACTION FUNCTIONS (Used by extract_main.py)
# ==========================================

def write_initial_sep_rows(ws_sep, records: list[dict]):
    """
    Clears sep_web_driver and writes header + one row per record:
      A: dttm_sep, B: sep_num, C: receipt_num, D: receipt_type,
      E: updated_dttm, F: status, G: note
    """
    rows =[
        [
            rec["dttm_sep"],            # original visit date/time
            rec["mrn"],                 # medical record number
            rec["sep_num"],             # SEP number
            rec["presc_id"],            # prescription id
            rec["receipt_num"],         # prescription number
            "Obat Kronis Blm Stabil",   # receipt_type (manual entry)
        ]
        for rec in records
    ]
    print(f"Writing {len(rows)} records to Google Sheet...")
    ws_sep.append_rows(rows)

# ==========================================
# 2. SUBMISSION FUNCTIONS (Used by submit_main.py)
#    Locking: G(Processing By), H(Processing Started)
#    Results: I(SubID), J(Time), K(Status), L(Note)
# ==========================================

def claim_row(ws, row_idx: int, ttl_seconds: int = 300, max_retries: int = 3, sleep: float = 0.4) -> bool:
    return _claim_row_generic(ws, row_idx, "G", "H", ttl_seconds, max_retries, sleep)

def commit_row_result(ws, row_idx: int, status: str, note: str, submission_id: str | None = None):
    ts = _now_iso()
    updates = []
    # Write Results (I, J, K, L)
    updates.append({"range": f"I{row_idx}", "values": [[submission_id or ""]]})
    updates.append({"range": f"J{row_idx}", "values": [[ts]]})
    updates.append({"range": f"K{row_idx}", "values": [[status or ""]]})
    updates.append({"range": f"L{row_idx}", "values": [[note or "-"]]})
    
    # Clear Locks (G, H)
    ws.batch_clear([f"G{row_idx}", f"H{row_idx}"])
    
    ws.batch_update(updates)

# ==========================================
# 3. REVISION FUNCTIONS (Used by revise_submit_main.py)
#    Locking: H(Processing By), I(Processing Started)
#    Results: J(SubID), K(Time), L(Status), M(Note)
# ==========================================

def claim_revise_row(ws, row_idx: int, ttl_seconds: int = 300, max_retries: int = 3, sleep: float = 0.4) -> bool:
    # Revise sheet structure is shifted +1 column compared to original
    return _claim_row_generic(ws, row_idx, "H", "I", ttl_seconds, max_retries, sleep)

def commit_revise_result(ws, row_idx: int, status: str, note: str, submission_id: str | None = None):
    ts = _now_iso()
    updates = []
    # Write Results (J, K, L, M)
    updates.append({"range": f"J{row_idx}", "values": [[submission_id or ""]]})
    updates.append({"range": f"K{row_idx}", "values": [[ts]]})
    updates.append({"range": f"L{row_idx}", "values": [[status or ""]]})
    updates.append({"range": f"M{row_idx}", "values": [[note or "-"]]})
    
    # Clear Locks (H, I)
    ws.batch_clear([f"H{row_idx}", f"I{row_idx}"])
    
    ws.batch_update(updates)

# ==========================================
# INTERNAL HELPER (Generic Locking)
# ==========================================

def _claim_row_generic(ws, row_idx: int, col_by: str, col_start: str, ttl_seconds: int, max_retries: int, sleep: float) -> bool:
    proc_by_cell = f"{col_by}{row_idx}"
    proc_started_cell = f"{col_start}{row_idx}"

    for attempt in range(max_retries):
        try:
            # 1. Check current owner
            current_owner = (ws.acell(proc_by_cell).value or "").strip()
            current_started = (ws.acell(proc_started_cell).value or "").strip()

            # 2. If empty, try to claim
            if not current_owner:
                batch = [
                    {"range": proc_by_cell, "values": [[HOSTNAME]]},
                    {"range": proc_started_cell, "values": [[_now_iso()]]},
                ]
                ws.batch_update(batch)
                time.sleep(0.25)
                # Verify we won the race
                confirm = (ws.acell(proc_by_cell).value or "").strip()
                if confirm == HOSTNAME:
                    return True
            
            # 3. If occupied, check TTL (stale lock)
            else:
                try:
                    started_dt = datetime.strptime(current_started, "%Y-%m-%dT%H:%M:%S")
                    if datetime.now() - started_dt > timedelta(seconds=ttl_seconds):
                        print(f"⚠️ Stealing stale lock from row {row_idx}...")
                        batch = [
                            {"range": proc_by_cell, "values": [[HOSTNAME]]},
                            {"range": proc_started_cell, "values": [[_now_iso()]]},
                        ]
                        ws.batch_update(batch)
                        time.sleep(0.25)
                        confirm = (ws.acell(proc_by_cell).value or "").strip()
                        if confirm == HOSTNAME:
                            return True
                except Exception:
                    # Date parse fail or other error -> treat as locked
                    pass
            
            time.sleep(sleep)
        except Exception:
            time.sleep(sleep)
    return False

# Legacy support
def update_sep_row(ws_sep, row_index: int, status: str, note: str):
    pass