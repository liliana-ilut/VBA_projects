'this code will copy the filtered table into a new workbook
Option Explicit

Sub Copy_Filtered_List_Workbook()

  ActiveSheet.AutoFilter.Range.Copy
  Workbooks.Add
  Range("A1").PasteSpecial
  
End Sub
----------------------------------------------------------------
'this code will copy the filtered table into a new worksheet
Sub Copy_Filtered_List_Spreadsheet()

  ActiveSheet.AutoFilter.Range.Copy
  Worksheets.Add
  Range("A1").PasteSpecial
  
End Sub
