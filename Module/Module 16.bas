Attribute VB_Name = "Module16"
Sub ShowSavingsAnalysis()
    Dim wsExpenses As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim categories As Object
    Dim cat As Variant
    Dim maxCategory As String
    Dim maxSpending As Double
    Dim msg As String

    Set wsExpenses = ThisWorkbook.Sheets("Expenses&Incomes")
    lastRow = wsExpenses.Cells(wsExpenses.Rows.Count, "A").End(xlUp).row

    Set categories = CreateObject("Scripting.Dictionary")

    For i = 4 To lastRow
        If wsExpenses.Cells(i, 3).Value <> "Income" Then
            cat = wsExpenses.Cells(i, 3).Value
            If categories.exists(cat) Then
                categories(cat) = categories(cat) + wsExpenses.Cells(i, 4).Value
            Else
                categories.Add cat, wsExpenses.Cells(i, 4).Value
            End If
        End If
    Next i

    maxSpending = 0
    For Each cat In categories.Keys
        If categories(cat) > maxSpending Then
            maxSpending = categories(cat)
            maxCategory = cat
        End If
    Next cat

    msg = "You are spending the most on " & maxCategory & ". Try to cut back on it."

    MsgBox msg
End Sub
