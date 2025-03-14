import win32com.client as win32
import pythoncom
from tkinter import filedialog
import tkinter as tk

def copy_cosco_actual():
    pythoncom.CoInitialize()
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    current_workbook = excel.ActiveWorkbook

    sheet_names = ["Block Fabrication"]

    # 지정된 시트 삭제
    for sheet in current_workbook.Sheets:
        if sheet.Name in sheet_names:
            sheet.Delete()

    # 파일 선택 대화 상자 열기
    root = tk.Tk()
    root.withdraw()
    source_file_path = filedialog.askopenfilename(
        title="Select a File",
        filetypes=[("Excel Files", "*.xls;*.xlsx;*.xlsm"), ("All Files", "*.*")]
    )

    if not source_file_path:
        print("No file selected.")
        return

    # 선택한 파일 열기
    source_workbook = excel.Workbooks.Open(source_file_path)

    excel.DisplayAlerts = False
    excel.ScreenUpdating = False

    for sheet_name in sheet_names:
        try:
            sheet = source_workbook.Sheets(sheet_name)
            
            # 값만 복사
            sheet.Cells.Copy()
            sheet.Cells.PasteSpecial(Paste=win32.constants.xlPasteValues)
            excel.CutCopyMode = False

            # 새 시트 생성
            target_sheet = current_workbook.Sheets.Add(After=current_workbook.Sheets(current_workbook.Sheets.Count))
            target_sheet.Name = sheet_name

            # 사용 범위 가져오기
            used_range = sheet.UsedRange

            # 서식 복사 (셀 크기, 병합 셀 등)
            used_range.Copy()
            target_sheet.Range(used_range.Address).PasteSpecial(Paste=win32.constants.xlPasteColumnWidths)
            target_sheet.Range(used_range.Address).PasteSpecial(Paste=win32.constants.xlPasteFormats)

            # 값 복사
            target_sheet.Range(used_range.Address).Value = used_range.Value

            # 병합된 셀 복사
            for merged_cell in used_range.MergeAreas:
                target_sheet.Range(merged_cell.Address).Merge()

            excel.CutCopyMode = False

        except Exception as e:
            print(f"Error processing sheet {sheet_name}: {str(e)}")

    excel.ScreenUpdating = True
    excel.DisplayAlerts = True

    # 선택한 파일 닫기
    source_workbook.Close(SaveChanges=False)

    current_workbook.Save()
    excel.Quit()
    pythoncom.CoUninitialize()

    print("Done")
    # compare_ranges() 함수 호출은 이 함수가 정의되어 있다고 가정합니다.
    # compare_ranges()

if __name__ == "__main__":
    copy_cosco_actual()