import os
import re
import shutil
from datetime import datetime, date

import tkinter as tk
from tkinter import ttk
from tkinter.filedialog import askopenfilename

import pandas as pd
from sqlalchemy import create_engine
import pyodbc  # noqa: F401  (ensure ODBC driver availability at runtime)

import win32com.client as win32
import pywintypes

# ---------------------- Ministry Selection (Dropdown with ALL) ---------------------- #
MINISTRY_LIST = [
    "MAG",
    "MCCSS",
    "MCURES",
    "MECP",
    "MEDJCT",
    "MEDU",
    "MEM",
    "MEPR",
    "MLTC",
    "MNEDG",
    "MNR",
    "MOH",
    "MOI",
    "MTCG",
    "MTO",
    "MTO-T",
    "SOLGEN",
]
CHOICES = ["ALL"] + MINISTRY_LIST

def select_ministry() -> str | None:
    """Modal dropdown to select a ministry (or ALL). Returns selection or None if canceled."""
    root = tk.Tk()
    root.withdraw()

    win = tk.Toplevel(root)
    win.title("Select a Ministry")
    win.resizable(False, False)
    win.grab_set()

    win.update_idletasks()
    w, h = 420, 140
    x = win.winfo_screenwidth() // 2 - w // 2
    y = win.winfo_screenheight() // 3 - h // 2
    win.geometry(f"{w}x{h}+{x}+{y}")

    tk.Label(win, text="Choose a ministry (or ALL):", anchor="w").pack(padx=16, pady=(16, 8), fill="x")

    selected = tk.StringVar(value="ALL")
    cb = ttk.Combobox(win, textvariable=selected, values=CHOICES, state="readonly", width=30)
    cb.pack(padx=16, fill="x")
    cb.focus_set()

    btn_frame = tk.Frame(win)
    btn_frame.pack(padx=16, pady=16, fill="x")

    result = {"value": None}
    def on_ok(event=None):
        val = selected.get().strip()
        if val in CHOICES:
            result["value"] = val
            win.destroy()
    def on_cancel(event=None):
        result["value"] = None
        win.destroy()

    ttk.Button(btn_frame, text="OK", command=on_ok).pack(side="right", padx=(8, 0))
    ttk.Button(btn_frame, text="Cancel", command=on_cancel).pack(side="right")

    win.bind("<Return>", on_ok)
    win.bind("<Escape>", on_cancel)
    win.protocol("WM_DELETE_WINDOW", on_cancel)

    root.wait_window(win)
    root.destroy()
    return result["value"]

selected_ministry = select_ministry()
if selected_ministry is None:
    print("No valid ministry was selected. Exiting the script.")
    raise SystemExit
print(f"The user selected: {selected_ministry}")

# ---------------------- File Setup ---------------------- #
template_path = askopenfilename(title="Please Select Empty Template")
if not template_path:
    print("No template selected. Exiting.")
    raise SystemExit

today = datetime.now()
date_text_ = today.strftime("%m_%d_%Y")

directory = os.path.dirname(template_path)
basename = os.path.basename(template_path)
_, extension = os.path.splitext(basename)

def safe_fname_token(s: str) -> str:
    # Replace any invalid filename characters (like /, \, :, *, ?, ", <, >, |) with underscores
    return "".join(["_" if c in r'\/:*?"<>|' else c for c in s])

# Single output file (either specific ministry or ALL)
file_token = "ALL" if selected_ministry == "ALL" else selected_ministry
output_name = f"{safe_fname_token(file_token)}_{date_text_}{extension}"
output_path = os.path.join(directory, output_name)
shutil.copy(template_path, output_path)

# ---------------------- SQL Connection ---------------------- #
server = "GSCVIKDCDBMSQ01"
database = "PipelineTracker"
driver = "ODBC Driver 17 for SQL Server"
connection_string = f"mssql+pyodbc://@{server}/{database}?driver={driver}&trusted_connection=yes"
engine = create_engine(connection_string)

sql_query = "SELECT * FROM Working_Table_Uploadtest_V2"
df_all = pd.read_sql(sql_query, engine)  # read once

# Build the source rows (single ministry or ALL in one DataFrame)
if selected_ministry == "ALL":
    source = df_all[df_all["Ministry"].isin(MINISTRY_LIST)].reset_index(drop=True)
else:
    source = df_all[df_all["Ministry"] == selected_ministry].reset_index(drop=True)

if source.empty:
    print("No rows found for the selected ministry selection.")
    # still produce an empty file with just the template; exit gracefully
    raise SystemExit

# ---------------------- Header Normalization ---------------------- #
def _norm(s: str) -> str:
    if s is None:
        return ""
    s = str(s)
    s = s.replace("\xa0", " ").replace("\r", " ").replace("\n", " ")
    s = s.strip().lower()
    s = re.sub(r'[\s_]+', ' ', s)       # collapse spaces/underscores
    s = re.sub(r'[^a-z0-9 ]+', '', s)   # drop punctuation
    return s

# ---------------------- Write to Excel (append ALL ministries into one file) ---------------------- #
excel = win32.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

try:
    workbook = excel.Workbooks.Open(output_path)
    try:
        sheet = workbook.Sheets("SOC Template")
        header_row = 1
        first_data_row = 6
        col_start = 1
        max_scan = 500

        # Discover Excel headers on header_row ONCE
        excel_headers = {}
        blanks = 0
        for c in range(col_start, col_start + max_scan):
            val = sheet.Cells(header_row, c).Value
            if val is None or str(val).strip() == "":
                blanks += 1
                if blanks >= 10:
                    break
                continue
            blanks = 0
            excel_headers[_norm(val)] = c

        manual_overrides = {}  # add mappings here if any DF col name doesn't match template header

        # Map DF columns to template columns (intersection only)
        df_to_excel_col = {
            col: excel_headers[_norm(manual_overrides.get(col, col))]
            for col in source.columns
            if _norm(manual_overrides.get(col, col)) in excel_headers
        }

        # Debug: make sure we found matches
        if not df_to_excel_col:
            print("[ERROR] No intersecting headers between DataFrame and template.")
            print("Sample DF cols (norm):", [ _norm(c) for c in list(source.columns)[:15] ])
            print("Sample template headers (norm):", list(excel_headers.keys())[:30])
            # don't save empty
            workbook.Close(SaveChanges=0)
            raise SystemExit

        # Decide write order: for ALL, preserve MINISTRY_LIST order; else just the single
        if selected_ministry == "ALL":
            write_order = MINISTRY_LIST
        else:
            write_order = [selected_ministry]

        # Start writing and append each ministry block
        next_excel_row = first_data_row
        total_written = 0

        for code in write_order:
            block = source[source["Ministry"] == code]
            if block.empty:
                continue
            for _, row in block.iterrows():
                for col_name, excel_col in df_to_excel_col.items():
                    if col_name not in row:
                        continue
                    value = row[col_name]
                    if pd.isna(value):
                        excel_value = ''
                    elif isinstance(value, pd.Timestamp):
                        excel_value = pywintypes.Time(value.to_pydatetime())
                    elif isinstance(value, date):
                        excel_value = pywintypes.Time(datetime(value.year, value.month, value.day))
                    else:
                        excel_value = value
                    sheet.Cells(next_excel_row, excel_col).Value = excel_value
                next_excel_row += 1
                total_written += 1

        workbook.Close(SaveChanges=1)
        print(f"Wrote {total_written} rows to {os.path.basename(output_path)}")

    except Exception:
        try:
            workbook.Close(SaveChanges=0)
        except:
            pass
        raise
finally:
    excel.Quit()
