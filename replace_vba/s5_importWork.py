import win32com.client as win32
import pythoncom

def import_work():
    pythoncom.CoInitialize()
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.ActiveWorkbook

    import_sheet = wb.Sheets("Import_Actual")
    source_all = wb.Sheets("Source_All")

    source_row = source_all.Range("O8:O728")
    source_col = source_all.Range("W6:EY6")

    ir, ic = 2, 1

    # Clear contents
    import_sheet.Rows("2:100000").ClearContents()
    import_sheet.Range("A1").Select()
    excel.CutCopyMode = False

    r = 0

    # BLOCK Status Loop
    for s_row in range(source_row.Row, source_row.Row + source_row.Rows.Count):
        for s_col in range(source_col.Column, source_col.Column + source_col.Columns.Count):
            p_start = source_all.Cells(6, s_col).Value[:7]
            stage_code = source_all.Cells(6, s_col).Value[-4:]
            stage_name = source_all.Cells(4, s_col).Value

            if source_all.Cells(s_row, 15).Value != "" and p_start == "P_Start":
                st_cell = source_all.Cells(s_row, s_col)

                import_sheet.Cells(ir + r, ic).Value = s_row
                import_sheet.Cells(ir + r, ic + 2).Value = source_all.Cells(s_row, 2).Value  # AREA
                import_sheet.Cells(ir + r, ic + 3).Value = source_all.Cells(s_row, 3).Value  # ZONE
                import_sheet.Cells(ir + r, ic + 4).Value = source_all.Cells(s_row, 4).Value  # MOD
                import_sheet.Cells(ir + r, ic + 5).Value = source_all.Cells(s_row, 4).Value  # LEVEL
                import_sheet.Cells(ir + r, ic + 6).Value = source_all.Cells(s_row, 5).Value  # NAME

                hg_2pe = source_all.Cells(s_row, 11).Value
                if len(str(hg_2pe)) <= 4:
                    hg_2pe = import_sheet.Cells(ir + r - 1, ic + 11).Value
                import_sheet.Cells(ir + r, ic + 11).Value = hg_2pe
                import_sheet.Cells(ir + r, ic + 12).Value = source_all.Cells(s_row, 11).Value  # H_2PE

                hg_1pe = source_all.Cells(s_row, 9).Value
                if len(str(hg_1pe)) <= 4:
                    hg_1pe = import_sheet.Cells(ir + r - 1, ic + 13).Value
                import_sheet.Cells(ir + r, ic + 13).Value = hg_1pe
                import_sheet.Cells(ir + r, ic + 14).Value = source_all.Cells(s_row, 9).Value  # H_1PE

                import_sheet.Cells(ir + r, ic + 19).Value = source_all.Cells(s_row, 15).Value  # BLOCK
                import_sheet.Cells(ir + r, ic + 20).Value = source_all.Cells(s_row, 16).Value  # SUBCON

                import_sheet.Cells(ir + r, ic + 21).Value = stage_code
                import_sheet.Cells(ir + r, ic + 22).Value = stage_name
                import_sheet.Cells(ir + r, ic + 23).Value = st_cell.Value  # PSD PLAN START DATE
                import_sheet.Cells(ir + r, ic + 24).Value = st_cell.Offset(0, 1).Value  # PFD PLAN FINISH DATE
                import_sheet.Cells(ir + r, ic + 25).Value = st_cell.Offset(0, 2).Value  # PPRO PLAN PROGRESS
                import_sheet.Cells(ir + r, ic + 26).Value = st_cell.Offset(0, 3).Value  # ASD ACTUAL START
                import_sheet.Cells(ir + r, ic + 27).Value = st_cell.Offset(0, 4).Value  # ASD ACTUAL FINISH
                import_sheet.Cells(ir + r, ic + 28).Value = st_cell.Offset(0, 5).Value  # APRO ACTUAL PRO

                r += 1

    excel.ScreenUpdating = True
    excel.DisplayAlerts = True

    wb.Save()
    excel.Quit()
    pythoncom.CoUninitialize()

    print("작업 완료")

if __name__ == "__main__":
    import_work()