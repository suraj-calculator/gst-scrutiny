"""
Extract Excel (.xlsx) files from every .zip file in a folder,
collect them into a single 'X' subfolder with clean (non-duplicated) names,
then delete all the original .zip files - leaving only the 'X' folder behind.

Usage:
    python run.py [folder]

`folder` is optional and defaults to the folder this script itself lives in
(so dropping run.py next to a batch of .zip files and running it with no
arguments just works, instead of relying on a hard-coded path).
"""

import os
import sys
import zipfile

# ---- CONFIG ----
DEST_FOLDER_NAME = "X"
EXCEL_EXTENSIONS = (".xlsx", ".xls", ".xlsm")
# ----------------


def extract_excel_from_zips(base_dir=None):
    if base_dir is None:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    base_dir = os.path.abspath(base_dir)
    dest_dir = os.path.join(base_dir, DEST_FOLDER_NAME)
    os.makedirs(dest_dir, exist_ok=True)

    if not os.path.isdir(base_dir):
        print(f"ERROR: Folder nahi mila -> {base_dir}")
        return

    zip_files = [f for f in os.listdir(base_dir) if f.lower().endswith(".zip")]

    if not zip_files:
        print("Is folder mein koi ZIP file nahi mili.")
        return

    moved_count = 0
    no_excel_zips = []
    used_names = set()

    for zip_name in sorted(zip_files):
        zip_path = os.path.join(base_dir, zip_name)

        try:
            with zipfile.ZipFile(zip_path, "r") as zf:
                excel_members = [
                    m for m in zf.namelist()
                    if m.lower().endswith(EXCEL_EXTENSIONS) and not m.endswith("/")
                ]

                if not excel_members:
                    no_excel_zips.append(zip_name)
                    print(f"Warning: '{zip_name}' ke andar koi Excel file nahi mili.")
                    continue

                for member in excel_members:
                    inner_filename = os.path.basename(member)

                    # Clean name = original filename inside the zip.
                    # If a name clash happens, add a numeric suffix to avoid overwrite.
                    final_name = inner_filename
                    if final_name in used_names:
                        name_part, ext = os.path.splitext(inner_filename)
                        counter = 2
                        while f"{name_part}_{counter}{ext}" in used_names:
                            counter += 1
                        final_name = f"{name_part}_{counter}{ext}"

                    used_names.add(final_name)
                    dest_path = os.path.join(dest_dir, final_name)

                    with zf.open(member) as source, open(dest_path, "wb") as target:
                        target.write(source.read())

                    moved_count += 1
                    print(f"Extracted: {zip_name} -> X/{final_name}")

        except zipfile.BadZipFile:
            print(f"ERROR: '{zip_name}' corrupt hai ya valid ZIP nahi hai, skip kiya.")

    # Ab sab zip files delete kar do
    deleted_count = 0
    for zip_name in zip_files:
        zip_path = os.path.join(base_dir, zip_name)
        try:
            os.remove(zip_path)
            deleted_count += 1
        except Exception as e:
            print(f"Warning: '{zip_name}' delete nahi ho payi -> {e}")

    print("\n----- SUMMARY -----")
    print(f"Excel files extracted : {moved_count}")
    print(f"ZIP files deleted      : {deleted_count}")
    if no_excel_zips:
        print(f"ZIPs with no Excel file: {no_excel_zips}")

    print(f"\n----- Files in '{DEST_FOLDER_NAME}' folder -----")
    x_files = sorted(os.listdir(dest_dir))
    if not x_files:
        print("(X folder khali hai)")
    else:
        for f in x_files:
            print(f)


if __name__ == "__main__":
    extract_excel_from_zips(sys.argv[1] if len(sys.argv) > 1 else None)
