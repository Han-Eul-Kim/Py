import win32com.client as win32
from tkinter import Tk, filedialog

excel = win32.Dispatch('Excel.Application')
excel.Visible = True

curfile = excel.Workbooks.Open("D:\Workspace\Py\haneultry.xlsx")


mywb =curfile.Workbooks.add()

myws=mywb.Worksheets('Sheet_FRI')

myws.Range('A1:A10').Value = 'HAPPY_FRI'

def file_selector(extension="*.xlsx"):
  root = Tk()
  root.withdraw()
  return filedialog.askopenfilename(
    title = "select a file including '0307' ",
    filetypes =[("Excel Files", extension)]   
  )