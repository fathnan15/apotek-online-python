# submit_main.py (only change: main() uses optional start/end args)
import time
import sys
import uuid
from apotek_runner  import init_apotek, submit_to_apotek, close_apotek
from sheets_handler import get_worksheet, read_all_records, update_sep_row, claim_row, commit_row_result
from config import WORKSHEET_NAME

def main():
    # Optional CLI args:
    # python submit_main.py [start_row] [end_row]
    # start_row/end_row are sheet row numbers (data rows start at 2)
    start_row = int(sys.argv[1]) if len(sys.argv) > 1 else 2
    end_row   = int(sys.argv[2]) if len(sys.argv) > 2 else None

    ws      = get_worksheet(WORKSHEET_NAME)
    records = read_all_records(ws)

    # If a concrete end_row is given, slice the records so we only iterate the requested rows.
    # records corresponds to sheet rows starting at 2 -> idx = row_index_in_sheet
    if end_row is not None:
        # compute python slice indices: records index 0 == sheet row 2
        slice_start = max(0, start_row - 2)
        slice_end = max(0, end_row - 1)  # exclusive end for python slice
        records = records[slice_start:slice_end]
        enumerated_start = start_row
    else:
        # start enumerating at start_row but include all records after that
        records = records[start_row - 2 :]
        enumerated_start = start_row

    init_apotek()
    for offset, row in enumerate(records):
        idx = enumerated_start + offset  # actual sheet row number
        # Skip already‐processed rows (idempotency): if submission_id or status present, skip
        if (row.get("submission_id", "") or "").strip() or (row.get("status", "") or "").strip():
            print(f"⏭ Row {idx} already done (submission_id/status present).")
            continue

        # Try to claim the row
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

            # Create submission_id for idempotency tracing
            submission_id = str(uuid.uuid4())
            commit_row_result(ws, idx, status, note, submission_id=submission_id)

            print(f"✅ Row {idx} updated: status={status}, note={note}")
            print(f"____________________________________________________________________")

        except Exception as e:
            # Ensure we commit an error and clear the claim
            try:
                commit_row_result(ws, idx, "error", str(e), submission_id=None)
            except Exception:
                pass
            print(f"❌ Row {idx} failed with exception: {e}")

    close_apotek()
    print("✅ All submissions complete.")

if __name__ == "__main__":
    main()
