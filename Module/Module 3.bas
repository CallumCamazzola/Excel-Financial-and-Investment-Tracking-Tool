Attribute VB_Name = "Module3"
Sub CommandButton2_Click()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim lastRow As Long
    
    Set ws = ThisWorkbook.Sheets("Expenses&Incomes")
    Set tbl = ws.ListObjects("ExpensesTable")

    If tbl.ListRows.Count > 0 Then
        ' Clear all data in the table
        tbl.DataBodyRange.ClearContents
        MsgBox "All data cleared successfully!", vbInformation
    Else
        MsgBox "No data to clear in the table.", vbExclamation
    End If
End Sub
