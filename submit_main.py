import sys
import uuid
import argparse
from apotek_runner  import init_apotek, submit_to_apotek, close_apotek
from sheets_handler import get_worksheet, read_all_records, update_sep_row, claim_row, commit_row_result
from config import WORKSHEET_NAME

def main(args):
    start_row = args.start
    end_row   = args.end
    port      = args.port

    ws      = get_worksheet(WORKSHEET_NAME)
    records = read_all_records(ws)

    # Slice records to match exact jurisdiction
    slice_start = max(0, start_row - 2)
    slice_end = max(0, end_row - 1)  
    records = records[slice_start:slice_end]
    enumerated_start = start_row

    # Pipe dynamic port into the runner
    cdp_url = f"http://127.0.0.1:{port}"
    init_apotek(cdp_endpoint=cdp_url)

    for offset, row in enumerate(records):
        idx = enumerated_start + offset 
        
        if (row.get("submission_id", "") or "").strip() or (row.get("status", "") or "").strip():
            print(f"⏭ Row {idx} already done (submission_id/status present).")
            continue

        claimed = claim_row(ws, row_idx=idx, ttl_seconds=300, max_retries=4)
        if not claimed:
            print(f"⏭ Row {idx} skipped (claimed by other worker).")
            continue

        try:
            rec_type = row.get("receipt_type", "").strip()
            if not rec_type:
                commit_row_result(ws, idx, "error", "missing receipt_type", submission_id=None)
                print(f"⚠️ Row {idx} missing receipt_type — marked error.")
                continue

            sep_num     = str(row.get("sep_num", "")).strip()
            receipt_num = str(row.get("receipt_num", "")).strip()

            print(f"▶️  Submitting row {idx}: SEP={sep_num}, Receipt={receipt_num}, Type={rec_type}")
            status, note = submit_to_apotek(sep_num, receipt_num, rec_type)

            submission_id = str(uuid.uuid4())
            commit_row_result(ws, idx, status, note, submission_id=submission_id)

            print(f"✅ Row {idx} updated: status={status}, note={note}")
            print(f"____________________________________________________________________")

        except Exception as e:
            try:
                commit_row_result(ws, idx, "error", str(e), submission_id=None)
            except Exception:
                pass
            print(f"❌ Row {idx} failed with exception: {e}")

    close_apotek()
    print("✅ All submissions complete.")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Submit to Apotek Worker")
    parser.add_argument("--port", type=str, required=True, help="Chrome CDP Port (e.g., 9222)")
    parser.add_argument("--start", type=int, required=True, help="Start row in the sheet")
    parser.add_argument("--end", type=int, required=True, help="End row in the sheet")
    
    args = parser.parse_args()
    main(args)