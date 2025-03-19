import win32com.client
import os
import datetime

def split_excel_sheets(file_path):  
  
  excel = win32com.client.Dispatch("Excel.Application")
  excel.Visible = False
  
  workbook = excel.Workbooks.Open(file_path)  
  today = datetime.datetime.now().strftime('%m%d')  
  
  for sheet in workbook.Sheets:
      new_workbook = excel.Workbooks.Add()
      sheet.Copy(Before=new_workbook.Sheets(1))
    
      new_file_path = os.path.join(os.path.dirname(file_path), f"{today}_{sheet.Name}.xlsm")
      new_workbook.SaveAs(new_file_path, FileFormat=52)
      new_workbook.Close()
  
  workbook.Close()  
  excel.Quit()

split_excel_sheets(r"D:\Workspace\Py\db.xlsm")