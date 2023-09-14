Option Explicit
Dim StartCell As Integer

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'use this code when you want to add an input box with only numeric values. By adding the loop will make the users insert only numeric values
Sub Input_Number_Only()
  Dim MyAnswer As String
  Do While IsNumeric(MyAnswer) = False
    MyAnswer = VBA.InputBox("Please Input Quantity" & vbNewLine & "It needs to be a numeric value!")
    If IsNumeric(MyAnswer) Then MsgBox "Well Done!"
  Loop
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Sub Simple_Do_Until_1()
  StartCell = 8
  Do Until StartCell = 14
    Range("B" & StartCell).Value = Range("A" & StartCell).Value + 10
    StartCell = StartCell + 1
  Loop
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub Simple_Do_Until_2()
StartCell = 8
  Do Until StartCell = 14
    Range("B" & StartCell).Value = Range("A" & StartCell).Value + 10
    StartCell = StartCell + 1
  Loop
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Sub Simple_Do_While()
StartCell = 8
  Do While Range("A" & StartCell).Value <> ""
    Range("C" & StartCell).Value = Range("A" & StartCell).Value + 10
    StartCell = StartCell + 1
  Loop
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Sub Simple_Do_Until_Conditional()
  StartCell = 8
  Do Until StartCell = 14
    If Range("A" & StartCell) = 0 Then Exit Do
    Range("D" & StartCell).Value = Range("A" & StartCell).Value + 10
    StartCell = StartCell + 1
  Loop
End Sub


