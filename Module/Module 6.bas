Attribute VB_Name = "Module6"
Sub ShowLineGraph()
    Dim ws As Worksheet
    Dim outputWs As Worksheet
    Dim tbl As ListObject
    Dim chartObj As ChartObject
    Dim incomeRange As Range
    Dim spendingRange As Range
    Dim dateRange As Range
    Dim categoryRange As Range
    Dim amountRange As Range
    Dim uniqueDates As Collection
    Dim incomeTotals() As Double
    Dim spendingTotals() As Double
    Dim dateArray() As Variant
    Dim lastRow As Long
    Dim cell As Range, i As Long
    Dim targetChartName As String
    
    ' Set target chart name
    targetChartName = "IncomeVsSpendingChart"
    
    ' Set worksheets and table
    Set ws = ThisWorkbook.Sheets("Expenses&Incomes")
    Set tbl = ws.ListObjects("ExpensesTable")
    
    ' Set ranges for Date, Amount, and Category columns
    Set dateRange = tbl.ListColumns("Date").DataBodyRange
    Set categoryRange = tbl.ListColumns("Category").DataBodyRange
    Set amountRange = tbl.ListColumns("Amount in $").DataBodyRange

    ' Gather unique dates
    Set uniqueDates = New Collection
    On Error Resume Next
    For Each cell In dateRange
        uniqueDates.Add cell.Value, CStr(cell.Value)
    Next cell
    On Error GoTo 0
    
    ' Convert uniqueDates collection to an array
    ReDim dateArray(1 To uniqueDates.Count)
    For i = 1 To uniqueDates.Count
        dateArray(i) = uniqueDates(i)
    Next i

    ReDim incomeTotals(1 To uniqueDates.Count)
    ReDim spendingTotals(1 To uniqueDates.Count)

    ' Calculate totals for income and spending by date
    For Each cell In dateRange
        For i = 1 To uniqueDates.Count
            If cell.Value = uniqueDates(i) Then
                If categoryRange.Cells(cell.row - dateRange.row + 1, 1).Value = "Income" Then
                    incomeTotals(i) = incomeTotals(i) + amountRange.Cells(cell.row - dateRange.row + 1, 1).Value
                Else
                    spendingTotals(i) = spendingTotals(i) + amountRange.Cells(cell.row - dateRange.row + 1, 1).Value
                End If
            End If
        Next i
    Next cell

    ' Set the "Output" sheet as outputWs
    Set outputWs = ThisWorkbook.Sheets("Output")

    ' Clear the chart created by this code
    For Each chartObj In outputWs.ChartObjects
        If chartObj.Name = targetChartName Then
            chartObj.Delete
            Exit For ' Exit loop after deleting the chart
        End If
    Next chartObj

    ' Create a new chart on the Output sheet
    Set chartObj = outputWs.ChartObjects.Add(Left:=outputWs.Range("D14").Left, Width:=375, Top:=outputWs.Range("D14").Top, Height:=225)
    chartObj.Name = targetChartName ' Assign the specific name to the new chart

    ' Add a series for spending
    With chartObj.Chart.SeriesCollection.NewSeries
        .XValues = dateArray ' Using the array of unique dates for the X-axis
        .Values = Application.Transpose(spendingTotals)
        .Name = "Spending"
    End With

    ' Add a series for income
    With chartObj.Chart.SeriesCollection.NewSeries
        .XValues = dateArray ' Using the array of unique dates for the X-axis
        .Values = Application.Transpose(incomeTotals)
        .Name = "Income"
    End With

    ' Format the chart
    chartObj.Chart.ChartType = xlLine
    chartObj.Chart.HasTitle = True
    chartObj.Chart.ChartTitle.Text = "Income vs Spending Over Time"
    chartObj.Chart.Axes(xlCategory, xlPrimary).HasTitle = True
    chartObj.Chart.Axes(xlCategory, xlPrimary).AxisTitle.Text = "Date"
    chartObj.Chart.Axes(xlValue, xlPrimary).HasTitle = True
    chartObj.Chart.Axes(xlValue, xlPrimary).AxisTitle.Text = "Amount"

    ' Activate the Output sheet
    outputWs.Activate
End Sub

