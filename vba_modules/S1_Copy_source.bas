Sub Copy_source()
    Dim sourceFilePath As String
    Dim sourceWorkbook As Workbook
    Dim currentWorkbook As Workbook
    Dim sheet As Worksheet
    Dim targetSheet As Worksheet
    Dim ws As Worksheet
    Dim shtName As Variant
    Dim sheetNames As Variant
    Dim usedRange As Range

    ' Set reference to current workbook

    Set currentWorkbook = ThisWorkbook

    ' 복사할 시트 이름 배열
    sheetNames = Array("Hull", "LQ", "Topside")

    ' 현재 워크북의 모든 시트를 반복하여 지정된 시트와 같은 시트의 내용을 모두 삭제

    Application.DisplayAlerts = False

    For Each ws In currentWorkbook.Sheets



        For Each shtName In sheetNames



            If ws.Name = shtName Then



                ws.Delete



                Exit For



            End If



        Next shtName



    Next ws



    Application.DisplayAlerts = True



    



    ' 파일 선택 대화 상자 열기 ------------------------------------------------------------------------------------------------------------------------



    sourceFilePath = Application.GetOpenFilename(FileFilter:="Excel Files (*.xls; *.xlsx), *.xls; *.xlsx", Title:="Select a File")



    



    ' 사용자가 파일을 선택하지 않고 취소했을 경우 -----------------------------------------



    If sourceFilePath = "False" Then



        MsgBox "No file selected.", vbExclamation



        Exit Sub



    End If



    



    Set currentWorkbook = ThisWorkbook



    



    ' 선택한 파일 열기 -------------------------------



    Set sourceWorkbook = Workbooks.Open(sourceFilePath)



    



    ' 지정된 시트 복사 (Values만) -------------------------------------------------------------------------



    Application.DisplayAlerts = False



    Application.ScreenUpdating = False



    



    For i = LBound(sheetNames) To UBound(sheetNames)



        On Error Resume Next



        Set sheet = sourceWorkbook.Sheets(sheetNames(i))



        If Not sheet Is Nothing Then



            ' 새 시트 생성



            Set targetSheet = currentWorkbook.Sheets.Add(After:=currentWorkbook.Sheets(currentWorkbook.Sheets.Count))



            targetSheet.Name = sheetNames(i)



            



            ' 사용되고 있는 범위만 가져오기



            Set usedRange = sheet.usedRange



            



            ' 서식 복사 (셀 크기, 병합 셀 등)



            usedRange.Copy



            targetSheet.Range(usedRange.Address).PasteSpecial xlPasteColumnWidths



            targetSheet.Range(usedRange.Address).PasteSpecial xlPasteFormats



            



            ' 값만 복사



            targetSheet.Range(usedRange.Address).Value = usedRange.Value



            



            ' 병합된 셀 복사



            For Each mergedCell In sheet.usedRange.MergeAreas



                targetSheet.Range(mergedCell.Address).Merge



            Next mergedCell



            



            Application.CutCopyMode = False



        End If



        Set sheet = Nothing



        Set targetSheet = Nothing



        On Error GoTo 0



    Next i



    



    Application.ScreenUpdating = True



    Application.DisplayAlerts = True



    



    ' 선택한 파일 닫기 ------------------------------------------------



    sourceWorkbook.Close SaveChanges:=False



    



    ' 정리



    Set sourceWorkbook = Nothing



    Set currentWorkbook = Nothing



    Set usedRange = Nothing



    



    Clear_GridLine



    



    MsgBox "Done", vbInformation



End Sub