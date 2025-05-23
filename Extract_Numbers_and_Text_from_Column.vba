Option Explicit
Const StartRow As Byte = 10
Dim LastRow As Long

' extract the text and numbers from a column and add them into 2 seperate columns
Sub For_Next_Loop_In_Text()
  Dim i As Long ' for looping inside each cell
  Dim MyValue As String
  Dim NumFound As Long
  Dim TextFound As String
  Dim r As Long 'for looping through rows
  
  LastRow = Range("A" & Rows.Count).End(xlUp).Row
  
  For r = StartRow To LastRow
    MyValue = Range("A" & r).Value
    For i = 1 To VBA.Len(MyValue)
      If IsNumeric(VBA.Mid(MyValue, i, 1)) Then
        NumFound = NumFound & Mid(MyValue, i, 1)
      ElseIf Not IsNumeric(Mid(MyValue, i, 1)) Then
        TextFound = TextFound & Mid(MyValue, i, 1)
      End If
    Next i
    Range("H" & r).Value = TextFound
    Range("I" & r).Value = NumFound
    NumFound = 0 'resetting the value to 0
    TextFound = "" 'resetting the value to 0
    
  Next r
End Sub

'clear the values from the columns you just created
Sub Clear_Values_For_Text_Loop()

  LastRow = Range("A" & Rows.Count).End(xlUp).Row
  Range("H" & StartRow, "I" & LastRow).ClearContents

End Sub
