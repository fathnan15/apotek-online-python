# sheets_handler.py

import gspread
from datetime import datetime
from google.oauth2.service_account import Credentials
from config import (
    SERVICE_ACCOUNT_PATH, SHEET_URL,
    WORKSHEET_NAME, KEMO_SHEET_NAME,
    SEP_SHEET_HEADERS
)

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

def get_worksheet(name: str):
    creds  = Credentials.from_service_account_file(SERVICE_ACCOUNT_PATH, scopes=SCOPES)
    client = gspread.authorize(creds)
    sheet  = client.open_by_url(SHEET_URL)
    return sheet.worksheet(name)

def write_initial_sep_rows(ws_sep, records: list[dict]):
    # Populate sep_web_driver with header + sep/receipt, empty rest
    rows = [SEP_SHEET_HEADERS] + [
        ["", rec["sep"], rec["receipt"], "", "", ""] for rec in records
    ]
    ws_sep.clear()
    ws_sep.update("A1", rows)

def find_kemo_row(ws_kemo, receipt: str):
    # Return 1-based row index in kemo_recp_num where column B == receipt
    all_rows = ws_kemo.get_all_records()
    for idx, row in enumerate(all_rows, start=2):
        if str(row.get("receipt_num","")).strip() == str(receipt).strip():
            return idx
    return None

def mark_kemo_inputed(ws_kemo, row_index: int, status: str = "inputed"):
    ws_kemo.update(f"C{row_index}", status)

def update_sep_row(ws_sep, row_index: int, rec_type: str, status: str, note: str):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    batch = [
        {"range": f"A{row_index}", "values": [[ts]]},
        {"range": f"D{row_index}", "values": [[rec_type]]},
        {"range": f"E{row_index}", "values": [[status]]},
        {"range": f"F{row_index}", "values": [[note or "-"]]}
    ]
    ws_sep.batch_update(batch)
