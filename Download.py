import os
import tkinter as tk
from tkinter import simpledialog, messagebox
from tkinter.filedialog import askopenfilename
import shutil
from sqlalchemy import create_engine
import pyodbc
import pandas as pd
import numpy as np
from datetime import datetime, date
import win32com.client as win32
import pywintypes

# ---------------------- Ministry Selection ---------------------- #
ministry = simpledialog.askstring(
    "Select a Ministry",
    "Enter Valid Ministry short hand: EDU, MAG, MCCSS, MECP, MLTC, MNR, MOH, MOI, MTO-H, MTO-T, SOLGEN"
)


def get_ministry():
    valid_options = {
        "EDU", "MAG", "MCCSS", "MECP", "MLTC", "MNR",
        "MOH", "MOI", "MTO-H", "MTO-T", "SOLGEN"
    }  # valid ministry options

    root = tk.Tk()
    root.withdraw()  # Hide the main window

    while True:
        global ministry

        if ministry is None:
            messagebox.showinfo("Input Cancelled", "No valid ministry was selected. Exiting.")
            root.destroy()
            return None

        if ministry in valid_options:
            print("Selected Ministry:", ministry)
            root.destroy()
            return ministry
        else:
            messagebox.showerror(
                "Invalid Input",
                "Please enter one valid ministry short hand: EDU, MAG, MCCSS, MECP, "
                "MLTC, MNR, MOH, MOI, MTO-H, MTO-T, SOLGEN"
            )
            ministry = simpledialog.askstring(
                "Select a Ministry",
                "Enter Valid Ministry short hand: EDU, MAG, MCCSS, MECP, MLTC, "
                "MNR, MOH, MOI, MTO-H, MTO-T, SOLGEN"
            )


selected_ministry = get_ministry()

if selected_ministry is not None:
    print(f"The user selected: {selected_ministry}")
else:
    print("No valid ministry was selected. Exiting the script.")

# ---------------------- File Setup ---------------------- #
template_path = askopenfilename(title="Please Select Empty Template")

today = datetime.now()
date_text_ = today.strftime("%m_%d_%Y")

directory = os.path.dirname(template_path)
basename = os.path.basename(template_path)
name, extention = os.path.splitext(basename)

output_name = f"{selected_ministry}_{date_text_}{extention}"
output_path = os.path.join(directory, output_name)

shutil.copy(template_path, output_path)

# ---------------------- SQL Connection ---------------------- #
server = "GSCVIKDCDBMSQ01"
database = "PipelineTracker"
driver = "ODBC Driver 17 for SQL Server"

connection_string = f"mssql+pyodbc://@{server}/{database}?driver={driver}&trusted_connection=yes"
engine = create_engine(connection_string)

sql_query = "SELECT * FROM Working_Table_Uploadtest_V2"
df = pd.read_sql(sql_query, engine)

# Filter by selected ministry
df = df[df['Ministry'] == ministry]

# Reset index for the loop
df = df.reset_index(drop=True)

# ---------------------- Data Cleaning ---------------------- #
columns_to_drop = [
    'Project Type (Social or Civil) ',
    'Estimated total Project Cost Variance ($M) (variance from approved cost)',
    "Risks due to TPC's age (Automated)",
    'Risks to Project Cost',
    'Impact (Risk to Project Cost)',
    'Likelihood (Alignment to Capital Refresh)',
    'Index (Risk : Capital Refresh)',
    'Likelihood - Progress by Spring 2026',
    'Index: Progress by 2026 to Capital Refresh',
    'Nb of days from Design Completion to EW',
    'Nb of days from Design Completion to RPF Issuance',
    'Nb of days from EW to RFP Issuance',
    'Nb of days from EW to Construction Start',
    'Nb of days from RFP Issuance to Construction Start',
    'Nb of days from Construction Start to Completion'
]

df = df.drop(columns=columns_to_drop)

# ---------------------- Format Columns ---------------------- #
date_columns = [
    'Date of Latest Estimated TPC',
    'Estimated Completion for Functional Program',
    'Estimated Completion for Environmental Assessment',
    'Estimated Completion Design',
    'If yes what is the start date for Early Works?',
    'If yes estimated DTC Completion Date',
    'RFP Issuance',
    'DPA Award (Progressive projects)',
    'Contract Award/ Construction Start',
    'Estimated Project Completion Date'
]


def convert_date_column(column):
    column = pd.to_datetime(column, errors='coerce')
    return column.dt.strftime('%m-%d-%Y')


for col in date_columns:
    df[col] = convert_date_column(df[col])

percentage_columns = [
    'Functional Program Readiness',
    'Environmental Assessment Readiness',
    'Design Readiness',
    'Construction Procurement/ RFP Readiness'
]


def convert_to_percentage(column):
    # Replace 'NULL' with NaN
    column.replace('NULL', np.nan, inplace=True)
    return column.apply(
        lambda x: f"{x * 100:.0f}%" if pd.notnull(x) and isinstance(x, (int, float)) else x
    )


for col in percentage_columns:
    df[col] = convert_to_percentage(df[col])

df.columns = range(df.shape[1])

# ---------------------- Write to Excel ---------------------- #
excel = win32.Dispatch("Excel.Application")
excel.Visible = False

workbook = excel.Workbooks.Open(output_path)

date_text = today.strftime("%m/%d/%Y")
sheet1 = workbook.Sheets("Data Validation")

cell = sheet1.Range("K4")
cell.Value = date_text

sheet = workbook.Sheets("Full Tracker Template")

skip_columns = [9, 15, 16, 17, 50, 51, 52, 53, 54, 55, 56, 57, 58, 59, 60, 61]
col_start = 1  # first column in Excel
row_start = 8  # first row in Excel

# Loop for adding data to excel
for i, row in df.iterrows():
    col_index = col_start
    for col_name, value in row.items():
        # Skip specified columns
        while col_index in skip_columns:
            col_index += 1

        # Set the cell value
        if pd.isnull(value):
            excel_value = ''
        elif isinstance(value, pd.Timestamp):
            excel_value = pywintypes.Time(value.to_pydatetime())
        elif isinstance(value, date):
            excel_value = pywintypes.Time(datetime(value.year, value.month, value.day))
        else:
            excel_value = value

        try:
            sheet.Cells(row_start + i, col_index).Value = excel_value
        except Exception as e:
            print(f"Error at row {i + row_start}, column {col_index} with value {value}: {e}")
            raise

        col_index += 1

# ---------------------- Add VBA Code ---------------------- #
sheet_module = sheet.CodeName
vba_project = workbook.VBProject
vba_module = vba_project.VBComponents(sheet_module)

vba_code = """
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim cell As Range
    On Error Resume Next
    For Each cell In Target
        If Not Intersect(cell, Me.Range("C:C, D:D, E:E, J:J, K:K, AI:AI, AH:AH, AP:AP")) Is Nothing Then
            If cell.Validation.Type <> 3 Then
                Application.Undo
                MsgBox "Copy and paste is not allowed in this column.", vbExclamation
            End If
        End If
    Next cell
    On Error GoTo 0
End Sub
"""

vba_module.CodeModule.AddFromString(vba_code)
workbook.Save()
workbook.Close()
excel.Quit()
