VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} StocksForm 
   Caption         =   "UserForm1"
   ClientHeight    =   6285
   ClientLeft      =   120
   ClientTop       =   460
   ClientWidth     =   8900.001
   OleObjectBlob   =   "StocksForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "StocksForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub UserForm_Initialize()
    Dim ws As Worksheet
    Dim stockRange As Range
    Set ws = ThisWorkbook.Sheets("Investments")
    
    ' Populate ComboBox with stock tickers
    Set stockRange = ws.Range("B5:B" & ws.Cells(ws.Rows.Count, "B").End(xlUp).row)
    Me.ComboBox1.List = Application.Transpose(stockRange.Value)
End Sub

Private Sub btnAddInvestment_Click()
    Dim ws As Worksheet
    Dim ticker As String
    Dim investAmount As Double
    Dim row As Long
    
    Set ws = ThisWorkbook.Sheets("Investments")
    
    If Me.ComboBox1.Value = "" Then
        MsgBox "Enter a stock"
        Exit Sub
    Else
        ticker = Me.ComboBox1.Value
    End If
    If IsNumeric(Me.TextBox1.Value) And Me.TextBox1.Value > 0 Then
        investAmount = CDbl(Me.TextBox1.Value)
    Else
        MsgBox "invalid investment amount"
        Exit Sub
    End If
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
    Dim wsHidden As Worksheet
    
    Set wsHidden = ThisWorkbook.Sheets("Backend Storage")
    Set ws = ThisWorkbook.Sheets("Investments")
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row
    
    ' Refresh the stock data (force update)
    ws.Calculate
    
    For i = 5 To lastRow
      percentChange = ws.Cells(i, 5).Value
      If percentChange * ws.Cells(i, 3).Value <> wsHidden.Cells(1, i).Value Then
        
        investmentValue = ws.Cells(i, 6).Value ' Investment Value column
        
        ' Calculate new gain/loss and accumulate it
        gainedLost = investmentValue * percentChange
        'making sure that if the stock percentchange hasnt updated it isn't compounding
        If gainedLost <> ws.Cells(i, 7) Then
            ws.Cells(i, 7).Value = ws.Cells(i, 7).Value + gainedLost
        End If
        
        ' Update Investment Value
        If investmentValue = 0 Then
            ws.Cells(i, 6).Value = ws.Cells(i, 3).Value ' Set to Amount Invested if 0
        Else
            ws.Cells(i, 6).Value = ws.Cells(i, 3).Value + ws.Cells(i, 7).Value
        End If
        wsHidden.Cells(1, i).Value = percentChange * ws.Cells(i, 3).Value
      End If
    Next i
End Sub


