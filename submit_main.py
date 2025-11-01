# submit_main.py

from apotek_runner  import init_apotek, submit_to_apotek, close_apotek
from sheets_handler import get_worksheet, read_all_records, update_sep_row, claim_row, commit_row_result
from config import WORKSHEET_NAME
import time
import sys
import uuid


def main():
    ws      = get_worksheet(WORKSHEET_NAME)
    records = read_all_records(ws)

    init_apotek()
    for idx, row in enumerate(records, start=2):
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