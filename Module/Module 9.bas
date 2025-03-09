Attribute VB_Name = "Module9"
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim tbl As ListObject
    Dim rng As Range
    Dim row As ListRow
    Dim ws As Worksheet
    
    MsgBox "running"
    Set ws = ThisWorkbook.Sheets("Financial Goals")
    
    Set tbl = ws.ListObjects("GoalTable")
    
    For Each row In tbl.ListRows
        If row.Range(1, 7).Value = 1 Then
            row.Range.EntireRow.Hidden = True
        Else
            row.Range.EntireRow.Hidden = False
        End If
    Next row
End Sub


