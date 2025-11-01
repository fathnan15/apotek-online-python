import gspread
import socket
import time
import uuid
from datetime import datetime, timedelta
from google.oauth2.service_account import Credentials
from config import (
    SERVICE_ACCOUNT_PATH,
    SHEET_URL,
    WORKSHEET_NAME,
    SEP_SHEET_HEADERS
)

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

def get_worksheet(name: str):
    """Authenticate and open the named worksheet."""
    creds  = Credentials.from_service_account_file(SERVICE_ACCOUNT_PATH, scopes=SCOPES)
    client = gspread.authorize(creds)
    sheet  = client.open_by_url(SHEET_URL)
    return sheet.worksheet(name)


def write_initial_sep_rows(ws_sep, records: list[dict]):
    """
    Clears sep_web_driver and writes header + one row per record:
      A: dttm_sep, B: sep_num, C: receipt_num, D: receipt_type,
      E: updated_dttm, F: status, G: note
    """
    rows =[
        [
            rec["dttm_sep"],    # original visit date/time
            rec["mrn"],         # medical record number
            rec["sep_num"],     # SEP number
            rec["receipt_num"], # prescription number
            "Obat Kronis Blm Stabil", # receipt_type (manual entry)
        ]
        for rec in records
    ]
    print(f"Writing {len(rows)} records to Google Sheet...")
    ws_sep.append_rows(rows)


def read_all_records(ws) -> list[dict]:
    values = ws.get_all_values()
    if not values:
        return []
    headers = values[0]
    data    = values[1:]
    records = [dict(zip(headers, row)) for row in data]
    return records


def update_sep_row(ws_sep, row_index: int, status: str, note: str):
    """
    Updates:
      E: updated_dttm,
      F: status,
      G: note
    """
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    batch = [
        {"range": f"F{row_index}", "values": [[ts]]},
        {"range": f"G{row_index}", "values": [[status]]},
        {"range": f"H{row_index}", "values": [[note or '-']]}
    ]
    ws_sep.batch_update(batch)


HOSTNAME = socket.gethostname()

def _now_iso():
    return datetime.now().strftime("%Y-%m-%dT%H:%M:%S")

def claim_row(ws, row_idx: int, ttl_seconds: int = 300, max_retries: int = 3, sleep: float = 0.4) -> bool:
    """
    Claim a row for processing using optimistic write+confirm.
    Sheet columns (your layout):
      F => processing_by
      G => processing_started
    Returns True if claim succeeded, False otherwise.
    """
    proc_by_cell = f"F{row_idx}"
    proc_started_cell = f"G{row_idx}"

    for attempt in range(max_retries):
        try:
            current_owner = (ws.acell(proc_by_cell).value or "").strip()
            current_started = (ws.acell(proc_started_cell).value or "").strip()

            if not current_owner:
                # write both cells in a single batch to reduce races
                batch = [
                    {"range": proc_by_cell, "values": [[HOSTNAME]]},
                    {"range": proc_started_cell, "values": [[_now_iso()]]},
                ]
                ws.batch_update(batch)
                time.sleep(0.25)
                confirm = (ws.acell(proc_by_cell).value or "").strip()
                if confirm == HOSTNAME:
                    return True
            else:
                # check TTL; attempt steal if stale
                try:
                    started_dt = datetime.strptime(current_started, "%Y-%m-%d %H:%M:%S")
                    if datetime.now() - started_dt > timedelta(seconds=ttl_seconds):
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
                    # parse fail / unexpected format — skip stealing this round
                    pass

            time.sleep(sleep)
        except Exception:
            # transient error — back off and retry
            time.sleep(sleep)
    return False


def release_row_claim(ws, row_idx: int):
    """
    Best-effort clear of claim columns (F, G).
    """
    proc_by_cell = f"F{row_idx}"
    proc_started_cell = f"G{row_idx}"
    try:
        batch = [
            {"range": proc_by_cell, "values": [[""]]},
            {"range": proc_started_cell, "values": [[""]]},
        ]
        ws.batch_update(batch)
    except Exception:
        pass


def commit_row_result(ws, row_idx: int, status: str, note: str, submission_id: str | None = None):
    """
    Commit result and clear claim.
    Target columns (per your header):
      H => submission_id
      I => updated_dttm
      J => status
      K => note
      F, G => cleared
    Uses batch_update to reduce RPCs.
    """
    ts = _now_iso()
    updates = []
    updates.append({"range": f"H{row_idx}", "values": [[submission_id or ""]]})
    updates.append({"range": f"I{row_idx}", "values": [[ts]]})
    updates.append({"range": f"J{row_idx}", "values": [[status or ""]]})
    updates.append({"range": f"K{row_idx}", "values": [[note or "-"]]})
    updates.append({"range": f"F{row_idx}", "values": [[""]]})
    updates.append({"range": f"G{row_idx}", "values": [[""]]})

    ws.batch_update(updates)