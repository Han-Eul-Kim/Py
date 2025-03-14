import win32com.client as win32
import tkinter as tk
from tkinter import filedialog #for window
import os
import time  # to check the process

def select_file(title, filetypes):
  root = tk.Tk() #init_Tkinter
  root.withdraw()
  return filedialog.askopenfilename(title=title, filetypes=filetypes)


xl = win32.Dispatch('Excel.Application')
xl.Visible = True

source_file = select_file("Select the source Excel File",[("Excel Files", "*.xlsx;*.xls;*.xlsm")])

if not source_file:
    print("No excel file selected")
    xl.Quit()
    exit()

source_wb = xl.Workbooks.Open(source_file)
souce_ws =source_wb.Sheets(1)

target_file = select_file("Select the target Excel File", [("Excel Files", "*.xlsx;*.xls;*.xlsm")])
if not target_file:
    print("No target file selected")
    source_wb.Close(SaveChanges=False)
    xl.Quit()
    exit()

target_wb = xl.Workbooks.Open(target_file)
target_ws = target_wb.Sheets(2)

target_ws.Cells.Clear()
time.sleep(1)

souce_ws.Copy(After = target_ws)
time.sleep(2)

target_ws.PasteSpecial(Paste=win32.constants.xlPasteAll,SkipBlanks = True)
time.sleep(2)

print("✅ Process completed successfully.")




target_ws.Save()
# target_ws.Close(SaveChanges=True)
# xl.Quit()





