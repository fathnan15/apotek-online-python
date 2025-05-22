# main.py

from sheets_handler    import (
    get_worksheet,
    write_initial_sep_rows,
    find_kemo_row,
    mark_kemo_inputed,
    update_sep_row
)
from playwright_runner import init_cdp, get_claim_records, close as close_sirs
from apotek_runner     import init_apotek, submit_to_apotek, close_apotek
from config            import WORKSHEET_NAME, KEMO_SHEET_NAME

def main():
    # 1) Extract from SIRS
    init_cdp()
    records = get_claim_records()
    close_sirs()

    # 2) Seed sheet
    ws_sep  = get_worksheet(WORKSHEET_NAME)
    ws_kemo = get_worksheet(KEMO_SHEET_NAME)
    write_initial_sep_rows(ws_sep, records)

    # 3) Submit to Apotek
    init_apotek()
    for idx, rec in enumerate(records, start=2):
        # Determine receipt type
        if find_kemo_row(ws_kemo, rec["receipt"]):
            rec_type = "Obat Kemoterapi"
        else:
            rec_type = "Obat Kronis Blm Stabil"

        # Submit into Apotek using that type
        status, note = submit_to_apotek(rec["sep"], rec["receipt"], rec_type)

        # Update Google Sheet
        update_sep_row(ws_sep, idx, rec_type, status, note)

        # Mark chemo sheet if needed
        krow = find_kemo_row(ws_kemo, rec["receipt"])
        if krow:
            mark_kemo_inputed(ws_kemo, krow)

        print(f"Row {idx}: SEP={rec['sep']} | type={rec_type} → {status}")

    close_apotek()
    print("✅ Full workflow complete.")

if __name__ == "__main__":
    main()
