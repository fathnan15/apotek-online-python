# manual_extract.py

from sirs_runner import init_sirs_manual, get_claim_records
from sheets_handler  import get_worksheet, write_initial_sep_rows
from config            import WORKSHEET_NAME

def main():
    # 1) Attach & wait for you to set filters
    init_sirs_manual()

    # 2) Scrape the SEP/NoResep list
    records = get_claim_records()

    # 3) Write into Google Sheet
    ws = get_worksheet(WORKSHEET_NAME)
    write_initial_sep_rows(ws, records)
    print(f"✅ Wrote {len(records)} records into your sheet.")

if __name__ == "__main__":
    main()
