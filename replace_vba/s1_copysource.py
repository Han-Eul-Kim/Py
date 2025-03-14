import os
import win32com.client as win32
import pythoncom
from tkinter import filedialog
import tkinter as tk
import logging
import time

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(message)s")

def copy_and_process_data():
    pythoncom.CoInitialize()
    
    # Excel 객체 생성
    try:
        excel = win32.DispatchEx("Excel.Application")
    except Exception as e:
        logging.error(f"Excel 객체 생성 실패: {str(e)}")
        return

    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False

    try:
        # 파일 선택 대화 상자 열기
        root = tk.Tk()
        root.withdraw()
        source_file_path = filedialog.askopenfilename(
            title="Select Source File",
            filetypes=[("Excel Files", "*.xls;*.xlsx;*.xlsm"), ("All Files", "*.*")]
        )

        if not source_file_path:
            logging.warning("❌ 파일이 선택되지 않았습니다.")
            return

        # 대상 파일 경로 설정
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        target_file_path = os.path.join(desktop, "20250307_Block Scuedule.xlsm")

        # 원본 파일 열기
        source_wb = excel.Workbooks.Open(source_file_path)
        time.sleep(1)  # 파일이 완전히 열릴 때까지 대기

        target_wb = excel.Workbooks.Open(target_file_path)
        time.sleep(1)  # 파일이 완전히 열릴 때까지 대기

        # 복사할 시트 목록
        sheet_names = ["Hull", "LQ", "Topside"]

        # 기존에 같은 이름의 시트 삭제
        for sheet in list(target_wb.Sheets):
            try:
                if sheet.Name in sheet_names:
                    sheet.Delete()
                    time.sleep(0.5)  # 시트 삭제 후 잠시 대기
            except Exception:
                logging.warning(f"⚠️ '{sheet.Name}' 시트 삭제 실패 (Excel 내부 오류)")

        # 시트 복사 및 처리
        for sheet_name in sheet_names:
            try:
                source_sheet = None
                for sheet in source_wb.Sheets:
                    if sheet.Name == sheet_name:
                        source_sheet = sheet
                        break
                
                if source_sheet is None:
                    logging.warning(f"⚠️ '{sheet_name}' 시트가 원본 파일에 없습니다. 새로 생성합니다.")
                    source_sheet = source_wb.Sheets.Add()
                    source_sheet.Name = sheet_name

                # 새 시트 추가 (시트가 없을 경우 대비)
                if target_wb.Sheets.Count == 0:
                    target_sheet = target_wb.Sheets.Add()
                else:
                    target_sheet = target_wb.Sheets.Add(After=target_wb.Sheets(target_wb.Sheets.Count))
                target_sheet.Name = sheet_name
                time.sleep(0.5)  # 시트 추가 후 잠시 대기

                # 데이터 복사
                if source_sheet.UsedRange.Rows.Count > 1:
                    used_range = source_sheet.UsedRange
                    target_sheet.Range(used_range.Address).Value = used_range.Value
                    target_sheet.Range(used_range.Address).ColumnWidth = used_range.ColumnWidth
                    target_sheet.Range(used_range.Address).RowHeight = used_range.RowHeight
                    target_sheet.Range(used_range.Address).NumberFormat = used_range.NumberFormat

                    # 병합된 셀 유지
                    if used_range.MergeCells:
                        for merged_cell in used_range.MergeAreas:
                            target_sheet.Range(merged_cell.Address).Merge()

                    target_sheet.Activate()
                    excel.ActiveWindow.DisplayGridlines = False

                    logging.info(f"✅ '{sheet_name}' 시트 처리 완료")
                else:
                    logging.warning(f"⚠️ '{sheet_name}' 시트에 데이터가 없습니다.")

                time.sleep(1)  # 각 시트 처리 후 잠시 대기

            except Exception as e:
                logging.error(f"❌ '{sheet_name}' 시트 처리 중 오류 발생: {str(e)}")

        # Import_Actual 시트 선택 (존재 여부 확인 후)
        if "Import_Actual" in [sheet.Name for sheet in target_wb.Sheets]:
            import_actual_sheet = target_wb.Sheets("Import_Actual")
            import_actual_sheet.Select()
            import_actual_sheet.Range("A1").Select()
        else:
            logging.warning("⚠️ 'Import_Actual' 시트가 대상 파일에 없습니다.")

        # 원본 파일 닫기
        source_wb.Close(SaveChanges=False)

        # 새로운 엑셀 파일 저장
        target_wb.SaveAs(target_file_path)
        logging.info(f"✅ 모든 데이터가 처리되었습니다! 저장 경로: {target_file_path}")

        return target_file_path  # 파일 경로 반환

    except Exception as e:
        logging.error(f"❌ 전체 프로세스 중 오류 발생: {str(e)}")
        return None

    finally:
        excel.ScreenUpdating = True
        excel.DisplayAlerts = True
        # Excel 애플리케이션을 종료하지 않음
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    copy_and_process_data()