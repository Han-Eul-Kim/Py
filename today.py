import win32com.client as win32
import pandas as pd

excel = win32.Dispatch('Excel.Application')
excel.Visible = True

db_wb = excel.Workbooks.Open(r"C:\Users\JGS\Desktop\DataBase0318.xlsx")

cb_ws = db_wb.Sheets("Combined")
db_ws = db_wb.Sheets("DataBase")

Stage_list = [list(stg) for stg in cb_ws.Range('A9:B17').Value]
lastRow = cb_ws.Cells(cb_ws.Rows.Count, "F").End(-4162).Row
blockRange = cb_ws.Range(cb_ws.Cells(8,7), cb_ws.Cells(lastRow,7))

pntRow = 2
for blk in blockRange:
  row = blk.Row  
  for stg in Stage_list:    
    if cb_ws.Cells(row, stg[0]).Value != '':
      arr = [
        pntRow-1,
        cb_ws.Cells(row, 6).Value, 
        cb_ws.Cells(row, 7).Value,
        cb_ws.Cells(row, 8).Value,
        cb_ws.Cells(7, stg[0]).Value, 
        cb_ws.Cells(4, stg[0]).Value,
        
        cb_ws.Cells(row, stg[0]+ 0).Value,
        cb_ws.Cells(row, stg[0]+ 1).Value,
        cb_ws.Cells(row, stg[0]+ 2).Value,
        cb_ws.Cells(row, stg[0]+ 3).Value,
        cb_ws.Cells(row, stg[0]+ 4).Value,
        cb_ws.Cells(row, stg[0]+ 5).Value
      ]
      db_ws.Range(db_ws.Cells(pntRow,1), db_ws.Cells(pntRow,12)).Value = [arr]
    pntRow = pntRow + 1
  
