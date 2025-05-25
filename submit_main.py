# submit_main.py

from sheets_handler import get_worksheet, read_all_records, update_sep_row
from apotek_runner  import init_apotek, submit_to_apotek, close_apotek
from config            import WORKSHEET_NAME

def main():
    # 1) Open the Google Sheet and read all existing rows
    ws      = get_worksheet(WORKSHEET_NAME)
    records = read_all_records(ws)

    # 2) Launch (attach to) the Apotek form in Chrome
    init_apotek()

    # 3) Iterate over each data row (1-based index with header in row 1)
    for idx, row in enumerate(records, start=2):
        rec_type = row.get("receipt_type", "").strip()
        if not rec_type:
            print(f"⚠️  Row {idx} missing receipt_type, skipping.")
            continue

        sep_num     = row["sep_num"]
        receipt_num = row["receipt_num"]

        print(f"▶️  Submitting row {idx}: SEP={sep_num}, Receipt={receipt_num}, Type={rec_type}")
        status, note = submit_to_apotek(sep_num, receipt_num, rec_type)

        # 4) Write back updated_dttm, status, and note into columns E, F, G
        update_sep_row(ws, idx, status, note)
        print(f"✅ Row {idx} updated: status={status}, note={note}")

    # 5) Tear down the Apotek Playwright session
    close_apotek()
    print("✅ All submissions complete.")

if __name__ == "__main__":
    main()
