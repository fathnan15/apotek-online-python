# submit_main.py

from sheets_handler import get_worksheet, read_all_records, update_sep_row
from apotek_runner  import init_apotek, submit_to_apotek, close_apotek
from config import WORKSHEET_NAME


def main():
    ws      = get_worksheet(WORKSHEET_NAME)
    records = read_all_records(ws)

    init_apotek()
    for idx, row in enumerate(records, start=2):
        # Skip already‐processed rows
        if row.get("status", "").strip():
            continue

        rec_type = row.get("receipt_type", "").strip()
        if not rec_type:
            print(f"⚠️  Row {idx} missing receipt_type; skipping.")
            continue

        # Cast to str _before_ strip to avoid the int-no-strip issue
        sep_num     = str(row.get("sep_num", "")).strip()
        receipt_num = str(row.get("receipt_num", "")).strip()

        print(f"▶️  Submitting row {idx}: SEP={sep_num}, Receipt={receipt_num}, Type={rec_type}")
        status, note = submit_to_apotek(sep_num, receipt_num, rec_type)

        update_sep_row(ws, idx, status, note)
        print(f"✅ Row {idx} updated: status={status}, note={note}")

    close_apotek()
    print("✅ All submissions complete.")

if __name__ == "__main__":
    main()