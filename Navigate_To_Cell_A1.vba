'this code will close all your worksheet and leave the cursor on Cell_A1 from the 1st Worksheet

Sub Workbook_beforeClose()
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        ws.Activate
        ws.Range("A1").Select
    Next ws
    ActiveWorkbook.Worksheets(1).Activate
End Sub
