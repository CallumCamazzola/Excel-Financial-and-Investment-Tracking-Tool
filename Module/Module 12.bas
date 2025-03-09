Attribute VB_Name = "Module12"
Sub RefreshDate()
    Dim ws As Worksheet
    Dim timeleft As String
    Dim lastRow As Long
    
    Set ws = ThisWorkbook.Sheets("Financial Goals")
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    For i = 4 To lastRow
         timeleft = DateDiff("d", Date, ws.Cells(i, 2).Value) & " days"
         ws.Cells(i, 3).Value = timeleft
    Next i
End Sub
