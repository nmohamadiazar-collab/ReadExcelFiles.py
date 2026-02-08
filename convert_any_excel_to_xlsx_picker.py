import os
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox

import win32com.client as win32


# ==========================
# SETTINGS
# ==========================
OUTPUT_FOLDER_NAME = "CONVERTED_XLSX"  # created next to this script


def ensure_folder(path: str) -> None:
    os.makedirs(path, exist_ok=True)


def make_safe_filename(name: str) -> str:
    # basic cleanup for Windows
    bad = '<>:"/\\|?*'
    for ch in bad:
        name = name.replace(ch, "_")
    return name.strip()


def convert_with_excel(excel_app, in_path: str, out_path: str) -> None:
    """
    Use Microsoft Excel to open and SaveAs .xlsx.
    FileFormat=51 is xlsx.
    """
    wb = excel_app.Workbooks.Open(in_path)
    try:
        wb.SaveAs(out_path, FileFormat=51)  # 51 = xlOpenXMLWorkbook (.xlsx)
    finally:
        wb.Close(False)


def main():
    # File picker
    root = tk.Tk()
    root.withdraw()

    files = filedialog.askopenfilenames(
        title="Select Excel file(s) to convert to .xlsx",
        filetypes=[
            ("Excel files", "*.xls *.xlsx *.xlsm"),
            ("All files", "*.*"),
        ],
    )

    if not files:
        print("No files selected. Exiting.")
        return

    base_dir = os.path.dirname(os.path.abspath(__file__))
    out_dir = os.path.join(base_dir, OUTPUT_FOLDER_NAME)
    ensure_folder(out_dir)

    # Start Excel automation
    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    converted = 0
    copied = 0
    failed = 0

    try:
        for in_path in files:
            in_path = os.path.abspath(in_path)
            fname = os.path.basename(in_path)
            ext = os.path.splitext(fname)[1].lower()

            out_name = make_safe_filename(os.path.splitext(fname)[0]) + ".xlsx"
            out_path = os.path.join(out_dir, out_name)

            print(f"\nProcessing: {fname}")

            try:
                # If already .xlsx, just copy (fastest, preserves exactly)
                if ext == ".xlsx":
                    shutil.copy2(in_path, out_path)
                    copied += 1
                    print(f"  Copied to: {out_path}")

                # If .xls or .xlsm, convert using Excel SaveAs
                elif ext in (".xls", ".xlsm"):
                    convert_with_excel(excel, in_path, out_path)
                    converted += 1
                    print(f"  Converted to: {out_path}")
                    if ext == ".xlsm":
                        print("  Note: macros are NOT kept in .xlsx output.")

                else:
                    failed += 1
                    print(f"  Skipped (not an Excel extension): {ext}")

            except Exception as e:
                failed += 1
                print(f"  FAILED: {e}")

    finally:
        excel.Quit()

    print("\n=== DONE ===")
    print("Copied (.xlsx):", copied)
    print("Converted (.xls/.xlsm -> .xlsx):", converted)
    print("Failed:", failed)
    print("Output folder:", out_dir)

    messagebox.showinfo(
        "Conversion complete",
        f"Copied (.xlsx): {copied}\nConverted (.xls/.xlsm): {converted}\nFailed: {failed}\n\nSaved to:\n{out_dir}"
    )


if __name__ == "__main__":
    main()
