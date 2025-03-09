VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InputForm 
   Caption         =   "UserForm1"
   ClientHeight    =   6285
   ClientLeft      =   120
   ClientTop       =   460
   ClientWidth     =   8900.001
   OleObjectBlob   =   "InputForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "InputForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim nextRow As Range
    Dim firstEmptyRow As Range
    Dim dayVal As Integer
    Dim monthVal As Integer
    Dim yearVal As Integer
    Dim combinedDate As String
    Dim i As Long
    Dim newRowNeeded As Boolean

    Set ws = ThisWorkbook.Sheets("Expenses&Incomes")
    Set tbl = ws.ListObjects("ExpensesTable")
    
    newRowNeeded = True
    For i = 1 To tbl.ListRows.Count
        If WorksheetFunction.CountA(tbl.ListRows(i).Range) = 0 Then
            Set firstEmptyRow = tbl.ListRows(i).Range
            newRowNeeded = False
            Exit For
        End If
    Next i

    'If no empty row is found, add a new row at the end of the table
    If newRowNeeded Then
        Set nextRow = tbl.ListRows.Add.Range
    Else
        Set nextRow = firstEmptyRow
    End If

    'required fields are not empty
    If txtDay.Value = "" Or txtMonth.Value = "" Or txtYear.Value = "" Or txtItem.Value = "" Or CmbCategory.Value = "" Or txtAmount.Value = "" Then
        MsgBox "Please fill in all fields.", vbExclamation
        Exit Sub
    End If

    If IsNumeric(txtDay.Value) And IsNumeric(txtMonth.Value) And IsNumeric(txtYear.Value) And IsNumeric(txtAmount.Value) Then
        dayVal = CInt(txtDay.Value)
        monthVal = CInt(txtMonth.Value)
        yearVal = CInt(txtYear.Value)
    Else
        MsgBox "Please enter numeric values", vbExclamation
        Exit Sub
    End If

    ' Combine day, month, and year into a single date format
    combinedDate = Format(DateSerial(yearVal, monthVal, dayVal), "mm/dd/yyyy")
    nextRow.Cells(1, 1).Value = combinedDate
    nextRow.Cells(1, 2).Value = txtItem.Value
    nextRow.Cells(1, 3).Value = CmbCategory.Value
    nextRow.Cells(1, 4).Value = txtAmount.Value

    txtDay.Value = ""
    txtMonth.Value = ""
    txtYear.Value = ""
    txtItem.Value = ""
    CmbCategory.Value = ""
    txtAmount.Value = ""

    MsgBox "Entry added successfully!", vbInformation
End Sub


Private Sub UserForm_Initialize()
    CmbCategory.AddItem "Shopping"
    CmbCategory.AddItem "Bills"
    CmbCategory.AddItem "Income"
    CmbCategory.AddItem "Entertainment"
    CmbCategory.AddItem "Other"
End Sub



