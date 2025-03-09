Attribute VB_Name = "Module15"
Sub ShowIncomeExpense()
    Dim wsExpenses As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim totalIncome As Double
    Dim totalExpense As Double
    Dim msg As String

    Set wsExpenses = ThisWorkbook.Sheets("Expenses&Incomes")
    lastRow = wsExpenses.Cells(wsExpenses.Rows.Count, "A").End(xlUp).row

    totalIncome = 0
    totalExpense = 0

    For i = 4 To lastRow
        If wsExpenses.Cells(i, 3).Value = "Income" Then
            totalIncome = totalIncome + wsExpenses.Cells(i, 4).Value
        Else
            totalExpense = totalExpense + wsExpenses.Cells(i, 4).Value
        End If
    Next i

    If totalIncome > totalExpense Then
        msg = "On track: Income is greater than expenses."
    Else
        msg = "Spend less: Income is less than expenses."
    End If

    MsgBox msg
End Sub
