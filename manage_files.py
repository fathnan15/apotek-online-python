import argparse
import logging
import sys
from pathlib import Path
import shutil

import pandas as pd


def setup_logging(verbose: bool):
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        level=level,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    )


def find_patient_folder(base_dir: Path, sep: str, mrn: str) -> Path:
    """Search recursively for a folder containing SEP and ending with MRN."""
    for d in base_dir.rglob("*"):
        if not d.is_dir():
            continue
        if sep in d.name and d.name.strip().endswith(mrn):
            return d
    return None


def resolve_source_path(files_dir: Path, raw_source: str, sep: str, mrn: str) -> Path:
    """
    Resolve source path: absolute, relative to files_dir, or recursive fallback.
    """
    candidate = Path(raw_source)
    if candidate.is_absolute() and candidate.exists():
        return candidate
    candidate = files_dir / raw_source
    if candidate.exists():
        return candidate
    return find_patient_folder(files_dir, sep, mrn)


def move_folder(src: Path, dst: Path, dry_run: bool) -> Path:
    """Move src folder into dst, returning the new folder path."""
    dst.mkdir(parents=True, exist_ok=True)
    new_loc = dst / src.name
    if dry_run:
        logging.info(f"DRY RUN: Would move '{src}' → '{new_loc}'")
    else:
        shutil.move(str(src), str(new_loc))
        logging.info(f"Moved '{src}' → '{new_loc}'")
    return new_loc


def copy_echo_file(echo_dir: Path, echo_filename: str, target_dir: Path, dry_run: bool):
    """Copy specified ECHO file into target_dir, respecting dry_run."""
    echo_path = echo_dir / echo_filename
    if not echo_path.exists():
        logging.warning(f"ECHO file '{echo_filename}' not found in {echo_dir}")
        return False
    if dry_run:
        logging.info(f"DRY RUN: Would copy ECHO '{echo_filename}' → '{target_dir}'")
    else:
        shutil.copy2(str(echo_path), str(target_dir))
        logging.info(f"Copied ECHO '{echo_filename}' → '{target_dir}'")
    return True


def main():
    parser = argparse.ArgumentParser(description="Move/organize patient folders, copy ECHO PDFs for all rows.")
    parser.add_argument("--excel", required=True, help="Excel file with operations list")
    parser.add_argument("--files-dir", required=True, help="Root folder containing patient subfolders")
    parser.add_argument("--echo-dir", required=True, help="Directory of ECHO PDF files")
    parser.add_argument("--update-status", action="store_true", help="Write updated status/source back to Excel")
    parser.add_argument("--dry-run", action="store_true", help="Show actions without performing them")
    parser.add_argument("--verbose", action="store_true", help="Enable detailed logging")
    parser.add_argument("--bupel", required=True, help="Use Bupel directory under files-dir")
    args = parser.parse_args()

    setup_logging(args.verbose)

    excel_path = Path(args.excel)
    files_dir  = Path(args.files_dir)
    echo_dir   = Path(args.echo_dir)
    bupel_dir = Path(args.bupel)
    dir  = files_dir / bupel_dir # Assuming this is the target directory for moved folders [always change to your actual target]

    # Check required paths
    for p in [excel_path, files_dir, echo_dir, dir]:
        if not p.exists():
            logging.error(f"Required path not found: {p}")
            sys.exit(1)

    # Load Excel
    df = pd.read_excel(excel_path, dtype=str).fillna("")

    for idx, row in df.iterrows():
        sep            = row.get("sep_num", "").strip()
        mrn            = row.get("mrn", "").strip()
        card           = row.get("card_num", "").strip()
        status         = row.get("status_cd", "").strip()
        raw_source     = row.get("source", "").strip()
        dest_sub       = row.get("destination", "").strip()
        echo_filename  = row.get("echo", "").strip()

        # Determine destination under 03.MARET
        if not dest_sub:
            logging.warning(f"Row {idx}: missing 'destination', skipping")
            target_dest = None
        else:
            target_dest = dir / dest_sub

        new_folder = None
        # Handle moves for relevant statuses
        if status in ("need_move", "not_found") and target_dest:
            if status == "need_move":
                src_path = resolve_source_path(files_dir, raw_source, sep, mrn)
            else:  # not_found
                found = find_patient_folder(files_dir, sep, mrn)
                if found:
                    df.at[idx, "source"] = str(found.relative_to(files_dir))
                src_path = found

            if not src_path or not src_path.exists():
                logging.warning(f"Row {idx}: source not found: '{raw_source}'")
            else:
                new_folder = move_folder(src_path, target_dest, args.dry_run)

            # Verify move and update status/source
            if new_folder and args.update_status and not args.dry_run:
                if new_folder.exists() and (status == "need_move" or status == "not_found"):
                    df.at[idx, "status_cd"] = "normal"
                    df.at[idx, "source"]    = str(new_folder.relative_to(files_dir))
                    logging.info(f"Row {idx}: move verified, status set to 'normal'")
                else:
                    logging.error(f"Row {idx}: move verification failed for '{new_folder}'")

        # Copy ECHO for any row with echo_filename
        if echo_filename:
            # determine target folder: moved folder or existing source
            target_folder = new_folder or resolve_source_path(files_dir, raw_source, sep, mrn)
            if not target_folder:
                logging.warning(f"Row {idx}: cannot resolve target folder for ECHO copy")
            else:
                success = copy_echo_file(echo_dir, echo_filename, target_folder, args.dry_run)
                if success:
                    logging.info(f"Row {idx}: ECHO '{echo_filename}' processed")
        else:
            logging.debug(f"Row {idx}: no ECHO file specified, skipping copy")

    # Write back updated Excel
    if args.update_status and not args.dry_run:
        backup = excel_path.with_suffix(".bak.xlsx")
        excel_path.rename(backup)
        df.to_excel(excel_path, index=False)
        logging.info(f"Excel updated; backup saved to '{backup}'")


if __name__ == "__main__":
    main()


'''
Example usage: 

python manage_files.py \
  --excel assets/documents/master/02.list_docs_februari_2026.xlsx \
  --files-dir assets/files/ \
  --bupel 02.Februari \
  --echo-dir assets/documents/echo \
  --update-status \
  --verbose \
  --dry-run
'''