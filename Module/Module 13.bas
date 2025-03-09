Attribute VB_Name = "Module13"
Sub GenerateOutputTable()
    Dim wsInvestments As Worksheet
    Dim wsOutput As Worksheet
    Dim tblOutput As ListObject
    Dim lastRow As Long
    Dim outputRow As ListRow
    Dim i As Long
    Dim j As Long
    
    Set wsInvestments = ThisWorkbook.Sheets("Investments")
    Set wsOutput = ThisWorkbook.Sheets("Output")

    Set tblOutput = wsOutput.ListObjects("Investments")
    
    
    ' clear the table contents
    If tblOutput.DataBodyRange Is Nothing Then
    Else
        tblOutput.DataBodyRange.ClearContents
    End If
    
    lastRow = wsInvestments.Cells(wsInvestments.Rows.Count, "B").End(xlUp).row
    j = 37
    For i = 5 To lastRow
        If wsInvestments.Cells(i, 3).Value > 0 Then
            j = j + 1
            wsOutput.Cells(j, 17).Value = wsInvestments.Cells(i, 2).Value ' Stock
            wsOutput.Cells(j, 18).Value = wsInvestments.Cells(i, 7).Value ' Amount Gained/Lost
        End If
    Next i
    Set chartRange = tblOutput.DataBodyRange
        ' Create a bar chart
        Set chartObj = wsOutput.ChartObjects.Add(Left:=wsOutput.Range("U39").Left, Top:=wsOutput.Range("U39").Top, Width:=400, Height:=300)
        With chartObj.Chart
            .SetSourceData Source:=chartRange
            .ChartType = xlBarClustered
            .HasTitle = True
            .ChartTitle.Text = "Bar Chart of Investments"
            
            ' Set the X and Y axis titles
            .Axes(xlCategory, xlPrimary).HasTitle = True
            .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Stock"
            .Axes(xlValue, xlPrimary).HasTitle = True
            .Axes(xlValue, xlPrimary).AxisTitle.Text = "Amount Gained/Lost"
            
            ' Customize chart appearance if needed
            .Axes(xlCategory).CategoryNames = tblOutput.ListColumns(1).DataBodyRange ' X-axis (Stock)
        End With
    
    ' Inform the user
    MsgBox "Output table and bar chart updated successfully!", vbInformation
End Sub

