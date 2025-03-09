Attribute VB_Name = "Module7"
Sub CreatePieChart()
    Dim wsOutput As Worksheet
    Dim wsExpenses As Worksheet
    Dim chartObj As ChartObject
    Dim lastRow As Long
    Dim categorySum As Object
    Dim cat As Variant
    Dim amount As Double
    Dim i As Long

    Set wsOutput = ThisWorkbook.Sheets("Output")
    Set wsExpenses = ThisWorkbook.Sheets("Expenses&Incomes")

    lastRow = wsExpenses.Cells(wsExpenses.Rows.Count, "A").End(xlUp).row

    ' Create a dictionary to sum amounts by category, excluding 'Income'
    Set categorySum = CreateObject("Scripting.Dictionary")
    For i = 2 To lastRow
        If wsExpenses.Cells(i, 3).Value <> "Income" Then
            cat = wsExpenses.Cells(i, 3).Value
            If IsNumeric(wsExpenses.Cells(i, 4).Value) Then
                amount = wsExpenses.Cells(i, 4).Value
                If categorySum.exists(cat) Then
                    categorySum(cat) = categorySum(cat) + amount
                Else
                    categorySum.Add cat, amount
                End If
            End If
        End If
    Next i

    Dim arrCategories() As Variant
    Dim arrValues() As Variant
    ReDim arrCategories(1 To categorySum.Count)
    ReDim arrValues(1 To categorySum.Count)

    Dim index As Integer
    index = 1
    For Each cat In categorySum.Keys
        arrCategories(index) = cat
        arrValues(index) = categorySum(cat)
        index = index + 1
    Next cat

    ' Add the chart to the output sheet
    Set chartObj = wsOutput.ChartObjects.Add(Left:=wsOutput.Range("R14").Left, Top:=wsOutput.Range("R14").Top, Width:=375, Height:=225)
    With chartObj.Chart
        .ChartType = xlPie
        .SetSourceData Source:=wsOutput.Range(wsOutput.Cells(1, 20), wsOutput.Cells(categorySum.Count, 21))
        .SeriesCollection.NewSeries
        .SeriesCollection(1).XValues = arrCategories
        .SeriesCollection(1).Values = arrValues
        .HasTitle = True
        .ChartTitle.Text = "Expense Breakdown by Category"
    End With
End Sub



