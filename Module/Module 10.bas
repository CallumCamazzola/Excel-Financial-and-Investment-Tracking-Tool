Attribute VB_Name = "Module10"
Private Sub UserForm_Initialize()
    Dim ws As Worksheet
    Dim stockRange As Range
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Adjust to your sheet name
    
    ' Populate ComboBox with stock tickers
    Set stockRange = ws.Range("B2:B" & ws.Cells(ws.Rows.Count, "B").End(xlUp).row)
    Me.ComboBox1.List = Application.Transpose(stockRange.Value)
End Sub

Private Sub btnAddInvestment_Click()
    Dim ws As Worksheet
    Dim ticker As String
    Dim investAmount As Double
    Dim row As Long
    
    Set ws = ThisWorkbook.Sheets("Sheet1")
    ticker = Me.ComboBox1.Value
    investAmount = CDbl(Me.TextBox1.Value)
    
    ' Find the row with the selected ticker
    row = Application.Match(ticker, ws.Columns("B"), 0)
    
    If Not IsError(row) Then
        ' Update Amount Invested
        ws.Cells(row, 3).Value = ws.Cells(row, 3).Value + investAmount
        
        ' Update Investment Value if it's 0
        If ws.Cells(row, 6).Value = 0 Then
            ws.Cells(row, 6).Value = ws.Cells(row, 3).Value
        End If
    Else
        MsgBox "Stock ticker not found.", vbExclamation
    End If
End Sub
Private Sub btnRefreshStockData_Click()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim percentChange As Double
    Dim investmentValue As Double
    Dim gainedLost As Double
    
    Set ws = ThisWorkbook.Sheets("Sheet1")
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row
    
    ' Refresh the stock data (force update)
    ws.Calculate
    
    For i = 2 To lastRow
        percentChange = ws.Cells(i, 5).Value ' Percent Change column
        investmentValue = ws.Cells(i, 6).Value ' Investment Value column
        
        ' Calculate new gain/loss and accumulate it
        gainedLost = investmentValue * percentChange
        ws.Cells(i, 7).Value = ws.Cells(i, 7).Value + gainedLost
        
        ' Update Investment Value
        If investmentValue = 0 Then
            ws.Cells(i, 6).Value = ws.Cells(i, 3).Value ' Set to Amount Invested if 0
        Else
            ws.Cells(i, 6).Value = ws.Cells(i, 3).Value + ws.Cells(i, 7).Value
        End If
    Next i
End Sub


