Option Explicit


Sub One_Find() ' find a match in a range and write the result in a cell next to it
  Dim CompID As Range
  Range("C3").ClearContents
  Set CompID = Range("A:A").Find(what:=Range("B3").Value, LookIn:=xlValues, lookat:=xlWhole)
  If Not CompID Is Nothing Then
    Range("C3").Value = CompID.Offset(, 4).Value
  Else
    MsgBox "Company not found!"
  End If
End Sub

Sub Many_Finds() ' find matches in a range and write the result in a column next to it
  Dim CompID As Range
  Dim i As Byte
  Dim FirstMatch As Variant
  Range("D3:D6").ClearContents
  i = 3
  'Dim start    'adding timer to check the speed of the code
  'start = Timer
  Set CompID = Range("A:A").Find(what:=Range("B3").Value, LookIn:=xlValues, lookat:=xlWhole)
  If Not CompID Is Nothing Then
    Range("D" & i).Value = CompID.Offset(, 4).Value
    FirstMatch = CompID.Address
    Do
      Set CompID = Range("A:A").FindNext(CompID)
      If CompID.Address = FirstMatch Then Exit Do
      i = i + 1
      Range("D" & i).Value = CompID.Offset(, 4).Value
    Loop
  Else
    MsgBox "Company not found!"
  End If
  'Debug.Print Timer - start
  'Application.Speech.Speak "Well Done!" & i - 2 & "Matches were found" 'adding speech when code finished running
End Sub
