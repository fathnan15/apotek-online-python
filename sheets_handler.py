import gspread
from datetime import datetime
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
    rows = [SEP_SHEET_HEADERS] + [
        [
            rec["dttm_sep"],    # original visit date/time
            rec["sep_num"],     # SEP number
            rec["receipt_num"], # prescription number
            "",                 # receipt_type (manual entry)
            "",                 # updated_dttm (to be filled on submission)
            "",                 # status
            ""                  # note
        ]
        for rec in records
    ]
    ws_sep.clear()
    ws_sep.update("A1", rows)


def read_all_records(ws) -> list[dict]:
    """Reads all rows into list of dicts keyed by headers."""
    return ws.get_all_records()


def update_sep_row(ws_sep, row_index: int, status: str, note: str):
    """
    Updates:
      E: updated_dttm,
      F: status,
      G: note
    """
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    batch = [
        {"range": f"E{row_index}", "values": [[ts]]},
        {"range": f"F{row_index}", "values": [[status]]},
        {"range": f"G{row_index}", "values": [[note or '-']]}
    ]
    ws_sep.batch_update(batch)
