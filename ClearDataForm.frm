VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ClearDataForm 
   Caption         =   "UserForm1"
   ClientHeight    =   6285
   ClientLeft      =   120
   ClientTop       =   460
   ClientWidth     =   8900.001
   OleObjectBlob   =   "ClearDataForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ClearDataForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim criteria As String
    Dim valueToClear As Variant
    Dim i As Long

    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Expenses&Incomes")

    ' Set the table (replace "MyTable" with the name of your table)
    Set tbl = ws.ListObjects("ExpensesTable")

    ' Get criteria and value from the userform
    criteria = ComboBox1.Value
    valueToClear = TextBox1.Value
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    ' Loop through table rows from bottom to top to avoid shifting issues
    For i = lastRow To 4 Step -1
        Select Case criteria
            Case "Date"
                If tbl.DataBodyRange(i, 1).Value = valueToClear Then
                    tbl.ListRows(i - 3).Delete
                End If
            Case "Item"
                If tbl.DataBodyRange(i, 2).Value = valueToClear Then
                    tbl.ListRows(i - 3).Delete
                End If
            Case "Category"
                If ws.Cells(i, 3).Value = valueToClear Then
                      tbl.ListRows(i - 3).Delete
                End If
        End Select
    Next i

    ' Unload the form after completion
    Unload Me
End Sub



Private Sub UserForm_Initialize()
    ComboBox1.AddItem "Item"
    ComboBox1.AddItem "Category"
End Sub


