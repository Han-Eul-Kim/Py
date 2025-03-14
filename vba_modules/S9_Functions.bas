



Function Clear_GridLine()



    Sheets(Array("Hull", "Hull_COSCO", "LQ", "Topside")).Select



    Sheets("Hull").Activate



    ActiveWindow.DisplayGridlines = False



    Sheets("Import_Actual").Select



    Range("A1").Select



End Function