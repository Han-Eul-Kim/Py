Sub CompareRanges()

    Dim HHI_org_header_range As Range
    Dim HHI_sou_header_range As Range
    Dim COSCO_org_header_range As Range
    Dim COSCO_sou_header_range As Range
    Dim org_cell As Range
    Dim source_cell As Range







    Dim isDifferent As Boolean







    Dim cellAddress As String







    Dim sheetNames As Variant







    Dim stName As Variant







    Dim Org_header As Range







    Dim sou_header As Range















    ' 복사할 시트 이름 배열







    sheetNames = Array("Hull", "Hull_COSCO", "LQ", "Topside")







    







    ' Header가 같은지 확인용







    Set HHI_org_header_range = ThisWorkbook.Sheets("Check_Source_Header").Range("B4:CN7")







    Set HHI_sou_header_range = ThisWorkbook.Sheets("Hull").Range("C4:CN7")







    







    ' Hull_COSCO 범위 설정







    Set COSCO_org_header_range = ThisWorkbook.Sheets("Check_Source_Header").Range("B16:BE19")







    Set COSCO_sou_header_range = ThisWorkbook.Sheets("Hull_COSCO").Range("B4:CO7")







    







    ' LQ 범위 설정







    Set LQ_org_header_range = ThisWorkbook.Sheets("Check_Source_Header").Range("B25:EG28")







    Set LQ_sou_header_range = ThisWorkbook.Sheets("LQ").Range("B4:EG7")







    







    ' Topside 범위 설정







    Set Topside_org_header_range = ThisWorkbook.Sheets("Check_Source_Header").Range("B34:DY37")







    Set Topside_sou_header_range = ThisWorkbook.Sheets("Topside").Range("B4:DY7")







    







    ' 두 범위의 셀을 하나씩 비교







    For Each stName In sheetNames







        If stName = "Hull_COSCO" Then







            Set Org_header = COSCO_org_header_range







            Set sou_header = COSCO_sou_header_range







        ElseIf stName = "LQ" Then







            Set Org_header = LQ_org_header_range







            Set sou_header = LQ_sou_header_range







        ElseIf stName = "Topside" Then







            Set Org_header = Topside_org_header_range







            Set sou_header = Topside_sou_header_range







        Else







            Set Org_header = HHI_org_header_range







            Set sou_header = HHI_sou_header_range







        End If















        ' 두 범위의 셀을 하나씩 비교







        For Each org_cell In Org_header







            ' source_header_range의 동일한 위치의 셀 가져오기







            Set source_cell = sou_header.Cells(org_cell.Row - Org_header.Row + 1, org_cell.Column - Org_header.Column + 1)







            







            ' 값이 다른 경우







            If org_cell.Value <> source_cell.Value Then







                cellAddress = org_cell.Address







                MsgBox "소스 Header 와 Form Header 값이 다릅니다!" & vbCrLf & _







                       "소스 시트 이름 : " & stName & vbCrLf & _







                       "셀 위치 : " & cellAddress & vbCrLf & _







                       "Import 폼 값 : " & org_cell.Value & vbCrLf & _







                       "소스 시트 값 : " & source_cell.Value, vbExclamation







                Exit Sub







            End If







        Next org_cell







    Next stName







    







    ' 값이 모두 같은 경우







    MsgBox "두 범위의 모든 값이 동일합니다.", vbInformation















End Sub













