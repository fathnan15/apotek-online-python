import os
import shutil
import pandas as pd

# === CONFIG ===
excel_path = "path/to/your_excel_file.xlsx"
root_dir   = "path/to/root"

# Load & prep
df = pd.read_excel(excel_path, dtype=str)
df["source"]   = ""     # will fill in
df["status_cd"] = df["status_cd"].fillna("")

# 1) Discover actual source folders for every sep_num
for idx, row in df.iterrows():
    sep = row["sep_num"]
    found = False

    for date_folder in os.listdir(root_dir):
        dfp = os.path.join(root_dir, date_folder)
        if not os.path.isdir(dfp):
            continue

        # look in each content folder name
        for content in os.listdir(dfp):
            if sep in content.replace(" ", ""):
                df.at[idx, "source"] = date_folder
                found = True
                break
        if found:
            break

    if not found:
        df.at[idx, "status_cd"] = "not_found"

# 2) Move only those marked need_move
to_move = df[df["status_cd"] == "need_move"]

for idx, row in to_move.iterrows():
    sep      = row["sep_num"]
    src_date = row["source"]
    dst_date = row["destination"]

    src_path = os.path.join(root_dir, src_date, 
                            next(f for f in os.listdir(os.path.join(root_dir, src_date))
                                 if sep in f.replace(" ", "")))
    dst_dir  = os.path.join(root_dir, dst_date)
    dst_path = os.path.join(dst_dir, os.path.basename(src_path))

    os.makedirs(dst_dir, exist_ok=True)
    shutil.move(src_path, dst_path)

    df.at[idx, "status_cd"] = "moved"
    print(f"✅ Moved {os.path.basename(src_path)}\n"
          f"    from  {src_date}\n"
          f"    to    {dst_date}")

# 3) Write your updated Excel back out
df.to_excel("path/to/your_excel_file.updated.xlsx", index=False)
