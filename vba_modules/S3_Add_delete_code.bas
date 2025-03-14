Sub Add_delete_code()



    Dim sheet As Worksheet



    Set sheet = ThisWorkbook.Sheets("Hull_COSCO")



    sheet.Activate



    



    sheet.Columns("C:AA").Select



    Selection.EntireColumn.Hidden = False



    



    Range("P8").Select



    With Selection.Interior



    .Color = 255



    End With



    ActiveCell.FormulaR1C1 = "DELETE"



    Range("P8").Select



    Selection.Copy



    Range(Selection, Selection.End(xlDown)).Select



    ActiveSheet.Paste



    Application.CutCopyMode = False



    



    Range("H8").Select



    With Selection.Interior



    .Color = 255



    End With



    ActiveCell.FormulaR1C1 = "DELETE"



    Range("H8").Select



    Selection.Copy



    Range(Selection, Selection.End(xlDown)).Select



    ActiveSheet.Paste



    Application.CutCopyMode = False



    



    MsgBox "DONE"







End Sub