import win32com.client as win32
import tkinter as tk
from tkinter import filedialog
import os
import random

ExcelApp = win32.Dispatch("Excel.Application") #progID for the dispath
ExcelApp.Visible = True

haneulWB = ExcelApp.Workbooks.Add()
haneulWS = haneulWB.Worksheets(1)

EXCRng1 = haneulWS.Range('C1:C10')
EXCRng1.Value = random.randint(1,100)

count = 0

for cell in EXCRng1:
    count += 1
    if count > 6:
        cell.Interior.Color = 255
        cell.Font.ColorIndex = 2
        cell.Font.Bold = True
    else: 
        cell.Interior.Color = 16711680
        cell.Font.ColorIndex =1
        cell.Font.Name = 'Roboto'
        
# add a sum formula(relative) to Range(D4)/ up for -, down for +
haneulWS.Range('D4').FormulaR1C1 = '=SUM(R[-3]C[-1]:R[2]C[-1])' 

haneulWS.Range('D5').Value = ExcelApp.WorksheetFunction.Sum(EXCRng1)

haneulWB.SaveAs('03_haneul_try')
haneulWB.Close()
ExcelApp.Quit()