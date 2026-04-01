# revise_submit_main.py
import sys
import uuid
import time
from apotek_runner  import init_apotek, submit_revision_to_apotek, close_apotek
from sheets_handler import get_worksheet, read_all_records, claim_revise_row, commit_revise_result
from config import REVISE_WORKSHEET_NAME

def main():
    # Usage: python revise_submit_main.py [start_row] [end_row]
    start_row = int(sys.argv[1]) if len(sys.argv) > 1 else 2
    end_row   = int(sys.argv[2]) if len(sys.argv) > 2 else None

    print(f"🔹 Opening Worksheet: {REVISE_WORKSHEET_NAME}")
    ws      = get_worksheet(REVISE_WORKSHEET_NAME)
    records = read_all_records(ws)

    # Slice records
    if end_row is not None:
        slice_start = max(0, start_row - 2)
        slice_end = max(0, end_row - 1)
        records = records[slice_start:slice_end]
        enumerated_start = start_row
    else:
        records = records[start_row - 2 :]
        enumerated_start = start_row

    init_apotek()
    
    for offset, row in enumerate(records):
        idx = enumerated_start + offset
        
        # Idempotency check (submission_id is column I in new layout)
        if (row.get("submission_id", "") or "").strip() or (row.get("status", "") or "").strip():
            print(f"⏭ Row {idx} already done.")
            continue

        # Claim Row (uses G/H in new layout)
        claimed = claim_revise_row(ws, row_idx=idx, ttl_seconds=300, max_retries=4)
        if not claimed:
            print(f"⏭ Row {idx} skipped (claimed by other worker).")
            continue

        try:
            rec_type = row.get("receipt_type", "").strip()
            
            # --- DATE HANDLING ---
            # Extract date from "sep_dttm" (Column A)
            # Row example: "01/02/2026"
            date_str = str(row.get("sep_dttm", "")).strip()
            
            if not rec_type:
                commit_revise_result(ws, idx, "error", "missing receipt_type", submission_id=None)
                print(f"⚠️ Row {idx} missing receipt_type.")
                continue
            
            if not date_str:
                commit_revise_result(ws, idx, "error", "missing date (sep_dttm)", submission_id=None)
                print(f"⚠️ Row {idx} missing sep_dttm (required for date).")
                continue

            sep_num     = str(row.get("sep_num", "")).strip()
            receipt_num = str(row.get("receipt_num", "")).strip()
            iteration_status = str(row.get("iterarion", "")).strip()

            print(f"▶️  Revise Row {idx}: SEP={sep_num}, Date={date_str}, Rec={receipt_num}")
            
            # Submit with Date Revision
            status, note = submit_revision_to_apotek(sep_num, receipt_num, rec_type, date_str, iteration_status)

            submission_id = str(uuid.uuid4())
            commit_revise_result(ws, idx, status, note, submission_id=submission_id)

            print(f"✅ Row {idx} updated: {status} | {note}")
            print(f"____________________________________________________________________")

        except Exception as e:
            try:
                commit_revise_result(ws, idx, "error", str(e), submission_id=None)
            except Exception:
                pass
            print(f"❌ Row {idx} failed: {e}")

    close_apotek()
    print("✅ All revisions complete.")

if __name__ == "__main__":
    main()