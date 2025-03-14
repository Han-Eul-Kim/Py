import win32com.client as win32

excel = win32.Dispatch('Excel.Application')
excel.Visible = True 

workbook = excel.Workbooks.Open(r"C:\path\to\file.xlsx")

sheet = workbook.Sheets(1) 

sheet.Range("A1:B2").Copy() 

sheet.Range("A3").PasteSpecial(Paste=win32.constants.xlPasteValues) 

workbook.Save()
workbook.Close(SaveChanges=True)
excel.Quit()