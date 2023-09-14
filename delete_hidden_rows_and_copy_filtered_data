'delete hidden rows or rows that are filtered out
Sub Delete_Hidden_Filtered_Rows()
  Dim r As Long
  
  LastRow = Range("A" & StartRow).CurrentRegion.Rows.Count + StartRow - 2
  For r = LastRow To StartRow Step -1
    If Rows(r).Hidden = True Then
      'Range("H" & r).Value = "X"   'use this first to ensure your code only deletes what you need. This line of code will add an "X" to hidden or filtered codes
      Rows(r).Delete 'you can comment this one out until you check that your code works
    End If
  Next r

End Sub


'copy filtered data into a new worksheet
Sub Copy_Filtered_List()

  ActiveSheet.AutoFilter.Range.Copy
  Worksheets.Add ' you can modify it if you want to add it to a new workbook instead of worksheet
  Range("A1").PasteSpecial
  
End Sub
