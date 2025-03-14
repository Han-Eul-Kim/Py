import openpyxl
import os

# relative ref.
wb = openpyxl.load_workbook("haneultry.xlsx") 

sheet1 = wb['HAN']
sheet2 = wb['EUL']

col_range = sheet1[2:3]

# print(os.environ)
print(os.environ.get('PATH'))

# #col_verticalValues
# for cols in sheet1.iter_cols():
#     for cell in cols:
#         print(cell.value, end=" \t")
#     print()