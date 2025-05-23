'this code will create an automatic Table of Contents for all your Worksheets in your open Workbook 

Sub Auto_Table_contents()
  Dim StartCell As Range
  Dim Sh As Worksheet
  Dim ShName As String
  Dim MsgConfirm As VBA.VbMsgBoxResult
  Dim EndCell As Range
  
 
  
  On Error Resume Next
  
  
  Set StartCell = Excel.Application.InputBox("Where do you want to insert the Table of Content?" _
  & vbNewLine & "Please select a cell: ", "Insert Table of Contents", , , , , , 8)
'make sue the table gets insterted in the first cell from that range
  
  If Err.Number = 424 Then Exit Sub
  On Error GoTo handle
  
  Set StartCell = StartCell.Cells(1, 1)
  Set EndCell = StartCell.Offset(Worksheets.Count - 2, 1)
  
  
  MsgConfirm = VBA.MsgBox("The values in cells:" & vbNewLine & StartCell.Address & " to " & EndCell.Address _
  & "will be overwritten." & vbNewLine & "Do u want to continue?", vbOKCancel + vbDefaultButton2, "Confirmation of overwritting cells")
  If MsgConfirm = vbCancel Then Exit Sub
  
  For Each Sh In Worksheets
    ShName = Sh.Name
    If ActiveSheet.Name <> ShName Then
      If Sh.Visible = xlSheetVisible Then
        ActiveSheet.Hyperlinks.Add anchor:=StartCell, Address:="", SubAddress:="'" & ShName & "'!A1", TextToDisplay:=ShName
        StartCell.Offset(0, 1).Value = Sh.Range("A1").Value
        Set StartCell = StartCell.Offset(1, 0)
      End If
    End If
  Next Sh
  Exit Sub
handle:
MsgBox "There is an error, please investigate!"

End Sub
