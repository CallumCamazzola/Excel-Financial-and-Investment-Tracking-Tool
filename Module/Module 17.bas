Attribute VB_Name = "Module17"
Sub StockAnalysis()
    Dim wsInvestments As Worksheet
    Dim tblInvestments As ListObject
    Dim maxAmount As Double
    Dim highestStock As String
    Dim lastRow As Long
    Dim i As Long
    Dim PercentTotal As Double
    Dim msg As String
    
    Set wsInvestments = ThisWorkbook.Sheets("Investments")
    
    Set tblInvestments = wsInvestments.ListObjects("Table5")
 
    maxAmount = -100
    highestStock = ""
    PercentTotal = 0
    msg = ""
    
    
    For i = 1 To tblInvestments.ListRows.Count
        With tblInvestments.ListRows(i).Range
            If IsNumeric(.Cells(7).Value) And .Cells(7).Value > maxAmount Then ' Amount Gained/Lost column (7th column)
                maxAmount = .Cells(7).Value
                highestStock = .Cells(2).Value ' Stock column (1st column)
            End If
        PercentTotal = PercentTotal + .Cells(5).Value
        End With
        
        
    Next i
    
    If PercentTotal > 0 Then
        msg = msg & " The Market had a good day today with the aggregate percent gain caluculated from all the tracked stocks being " & Format(PercentTotal * 100, "0.00") & "%" & vbCrLf
    Else
        msg = msg & " The Market had a down day today with the aggregate percent gain caluculated from all the tracked stocks being " & Format(PercentTotal * 100, "0.00") & "%" & vbCrLf
    End If
    msg = msg & "your best preforming stock is: " & highestStock & vbNewLine & _
            "with a profit of: " & Format(maxAmount, "Currency")
    MsgBox msg
End Sub


