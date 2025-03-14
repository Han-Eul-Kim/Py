Sub Combined_Allsource()



    Dim sh_Hull As Worksheet



    Dim sh_COSCO As Worksheet



    Dim sh_LQ As Worksheet



    Dim sh_Topside As Worksheet



    Dim org_cell As Range



    Dim source_cell As Range







    Dim stName As Variant



    Dim Org_header As Range



    Dim sou_header As Range







    ' 복사할 시트 이름 배열



    sheetNames = Array("Hull", "Hull_COSCO", "LQ", "Topside")



    



    '아래 셋팅값은 Source 원본시트가 변동이 있을때 수정 하시오----------------------



    Set importSheet = ThisWorkbook.Sheets("Import_Actual")



    Set Source_All = ThisWorkbook.Sheets("Source_All")



    



    '아래 셋팅값은 HHI 오는 Source Data 엑셀시트에 값이 변동될시 수정해야함



    Set Sht_Hull = ThisWorkbook.Sheets("Hull")



    Set Sht_COSCO = ThisWorkbook.Sheets("Hull_COSCO")



    Set Sht_LQ = ThisWorkbook.Sheets("LQ")



    Set Sht_Topside = ThisWorkbook.Sheets("Topside")



  '-------------------------------------------------------------------------------



  ' CLEAR



    Set currentWorkbook = Source_All



    ' 'Import' 시트에서 2번 행 아래의 모든 값을 삭제----------------------------







    currentWorkbook.Activate



    currentWorkbook.Rows("8:100000").Select



      Selection.Clear



    currentWorkbook.Range("A1").Select



     Application.CutCopyMode = False







    ' HHI Hull 범위 설정==========================================================







        HHI_lastRow = Sht_Hull.Cells(Sht_Hull.Rows.Count, "P").End(xlUp).Row



    Set HHI_source_row = Sht_Hull.Range("P8:P" & HHI_lastRow)



    Set HHI_source_cow = Sht_Hull.Range("P6:CO6")







    'Hull_COSCO 범위 설정



      Cosco_lastRow = Sht_COSCO.Cells(Sht_COSCO.Rows.Count, "O").End(xlUp).Row



    Set COSCO_source_row = Sht_COSCO.Range("O8:O" & Cosco_lastRow)



    Set COSCO_source_cow = Sht_COSCO.Range("B6:BE6")











    ' LQ 범위 설정



        LQ_lastRow = Sht_LQ.Cells(Sht_LQ.Rows.Count, "K").End(xlUp).Row



    Set LQ_source_row = Sht_LQ.Range("K8:K" & LQ_lastRow)



    Set LQ_source_cow = Sht_LQ.Range("K6:EG6")



    



    ' Topside 범위 설정



        Topside_lastRow = Sht_Topside.Cells(Sht_LQ.Rows.Count, "F").End(xlUp).Row



    Set Topside_source_row = Sht_Topside.Range("F8:F" & LQ_lastRow)



    Set Topside_source_cow = Sht_Topside.Range("F6:DY6")











    'Hull All Sheet 복사



    Sht_Hull.Range("C8:CO" & HHI_lastRow).Copy Destination:=ThisWorkbook.Sheets("Source_All").Range("B8")







    'LQ 시트복사-------------------------------------------------------------------------------------------



    'Source_All Sheet 에 마지막열 찾기



    LastRow = Source_All.Cells(Source_All.Rows.Count, "O").End(xlUp).Row



    Set lastCell = Source_All.Cells(LastRow + 1, "B")







    Sht_LQ.Range("B8:C" & LQ_lastRow).Copy Destination:=Source_All.Cells(LastRow + 1, "B") 'Area Block



    Sht_LQ.Range("D8:J" & LQ_lastRow).Copy Destination:=Source_All.Cells(LastRow + 1, "F") '1st Information Block



    Sht_LQ.Range("K8:Y" & LQ_lastRow).Copy Destination:=Source_All.Cells(LastRow + 1, "O") '2nd Information Block up to Cutting



    Sht_LQ.Range("AU8:BA" & LQ_lastRow).Copy Destination:=Source_All.Cells(LastRow + 1, "AD") 'Fab



    Sht_LQ.Range("BP8:BV" & LQ_lastRow).Copy Destination:=Source_All.Cells(LastRow + 1, "AK") 'Main Assembly



    Sht_LQ.Range("CK8:DE" & LQ_lastRow).Copy Destination:=Source_All.Cells(LastRow + 1, "AR") 'Install of Pre-Outfitting/1PE/Painting



    Sht_LQ.Range("DM8:DZ" & LQ_lastRow).Copy Destination:=Source_All.Cells(LastRow + 1, "BM") '2nd PE/ Install of Outfitting



    Sht_LQ.Range("EA8:EG" & LQ_lastRow).Copy Destination:=Source_All.Cells(LastRow + 1, "CH") 'Setting



    



    Sht_LQ.Range("Z8:AF" & LQ_lastRow).Copy Destination:=Source_All.Cells(LastRow + 1, "CV") 'Cutting of Wall



    Sht_LQ.Range("AG8:AM" & LQ_lastRow).Copy Destination:=Source_All.Cells(LastRow + 1, "DC") 'Cutting of Secondary Beam



    Sht_LQ.Range("AN8:AT" & LQ_lastRow).Copy Destination:=Source_All.Cells(LastRow + 1, "DJ") 'Cutting of Leg/Brace



    



    Sht_LQ.Range("BB8:BH" & LQ_lastRow).Copy Destination:=Source_All.Cells(LastRow + 1, "DQ") 'Fabrication  of Wall



    Sht_LQ.Range("BI8:BO" & LQ_lastRow).Copy Destination:=Source_All.Cells(LastRow + 1, "DX") 'Fabrication  of Leg/Brace



    Sht_LQ.Range("BW8:CC" & LQ_lastRow).Copy Destination:=Source_All.Cells(LastRow + 1, "EE") 'Leg Assembly



    Sht_LQ.Range("CD8:CJ" & LQ_lastRow).Copy Destination:=Source_All.Cells(LastRow + 1, "EL") 'Wall Installation







    Sht_LQ.Range("DF8:DL" & LQ_lastRow).Copy Destination:=Source_All.Cells(LastRow + 1, "ES") 'Lag Painting







    'Hull COSCO 시트복사-------------------------------------------------------------------------------------------



    'Source_All Sheet 에 마지막열 찾기



    LastRow = Source_All.Cells(Source_All.Rows.Count, "O").End(xlUp).Row



    Set lastCell = Source_All.Cells(LastRow + 1, "B")







    Sht_COSCO.Range("B8:V" & Cosco_lastRow).Copy Destination:=Source_All.Cells(LastRow + 1, "B") 'information block



    Sht_COSCO.Range("W8:AJ" & Cosco_lastRow).Copy Destination:=Source_All.Cells(LastRow + 1, "AK") 'Assembly and Pre-Outfitting



    Sht_COSCO.Range("AK8:AQ" & Cosco_lastRow).Copy Destination:=Source_All.Cells(LastRow + 1, "BF") 'Block Painting



    Sht_COSCO.Range("AR8:AX" & Cosco_lastRow).Copy Destination:=Source_All.Cells(LastRow + 1, "AY") '1st PE



    Sht_COSCO.Range("AY8:BE" & Cosco_lastRow).Copy Destination:=Source_All.Cells(LastRow + 1, "CO") 'Touch up Painting







    'Topside 시트복사-------------------------------------------------------------------------------------------



    'Source_All Sheet 에 마지막열 찾기



    LastRow = Source_All.Cells(Source_All.Rows.Count, "O").End(xlUp).Row



    Set lastCell = Source_All.Cells(LastRow + 1, "B")







    Sht_Topside.Range("B8:L" & Topside_lastRow).Copy Destination:=Source_All.Cells(LastRow + 1, "B") 'information block // PE Assign column 추가 안됐음



    Sht_Topside.Range("F8:F" & Topside_lastRow).Copy Destination:=Source_All.Cells(LastRow + 1, "O") 'BLOCK NO



    Sht_Topside.Range("H8:H" & Topside_lastRow).Copy Destination:=Source_All.Cells(LastRow + 1, "P") 'BLOCK NO



    Sht_Topside.Range("Y8:AE" & Topside_lastRow).Copy Destination:=Source_All.Cells(LastRow + 1, "W") 'Cutting



    Sht_Topside.Range("AU8:BA" & Topside_lastRow).Copy Destination:=Source_All.Cells(LastRow + 1, "AD") 'Fabrication



    Sht_Topside.Range("BN8:BT" & Topside_lastRow).Copy Destination:=Source_All.Cells(LastRow + 1, "AK") 'ASSEMBLY



    Sht_Topside.Range("CD8:CX" & Topside_lastRow).Copy Destination:=Source_All.Cells(LastRow + 1, "AR") 'PREOUTFITTING, 1PE, PAINTING



    Sht_Topside.Range("DE8:DR" & Topside_lastRow).Copy Destination:=Source_All.Cells(LastRow + 1, "BM") '2PE, OUTFITTING



    Sht_Topside.Range("DS8:DY" & Topside_lastRow).Copy Destination:=Source_All.Cells(LastRow + 1, "CH") 'SETTING







MsgBox "작업 완료"



End Sub