# test_apotek_and_sheet.py

from sheets_handler import get_worksheet, read_all_records, update_sep_row
from apotek_runner  import init_apotek, submit_to_apotek, close_apotek
from config            import WORKSHEET_NAME

def main():
    # 1) Grab your sheet & the first record
    ws      = get_worksheet(WORKSHEET_NAME)
    records = read_all_records(ws)
    if not records:
        print("❌ No records found in the sheet.")
        return

    # we'll test on row 2 (first data row)
    row_index = 2
    row       = records[0]
    sep       = row["sep_num"]
    receipt   = row["receipt_num"]
    rec_type  = row["receipt_type"].strip()

    if not rec_type:
        print(f"⚠️  Row {row_index} has no receipt_type—please fill column D and retry.")
        return

    print(f"▶️  Testing row {row_index}: SEP={sep}, Receipt={receipt}, Type={rec_type}")

    # 2) Attach to Apotek and submit
    init_apotek()
    status, note = submit_to_apotek(sep, receipt, rec_type)
    close_apotek()

    # 3) Write back into columns E–G
    update_sep_row(ws, row_index, status, note)
    print(f"✅ Updated row {row_index}: status={status}, note={note}")

if __name__ == "__main__":
    main()
