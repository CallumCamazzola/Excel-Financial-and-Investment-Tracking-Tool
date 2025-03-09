VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EnterGoals 
   Caption         =   "UserForm1"
   ClientHeight    =   6285
   ClientLeft      =   120
   ClientTop       =   460
   ClientWidth     =   8900.001
   OleObjectBlob   =   "EnterGoalsForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "EnterGoals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton4_Click()
    Dim wsGoals As Worksheet
    Dim wsBalance As Worksheet
    Dim lastRow As Long
    Dim goalDate As Date
    Dim dayVal As Integer
    Dim monthVal As Integer
    Dim yearVal As Integer
    Dim InitialAmount As Double
    Dim currentBalance As Double
    Dim amountRemaining As Double
    Dim timeleft As String
    Dim i As Long
    Dim firstEmptyRow As Range
    Dim newRowNeeded As Boolean
    
    Set wsGoals = ThisWorkbook.Sheets("Financial Goals")
    Set wsBalance = ThisWorkbook.Sheets("Expenses&Incomes")

    ' fields are filled
    If GoalName.Value = "" Or Day.Value = "" Or Month.Value = "" Or Year.Value = "" Then
        MsgBox "Please fill in all fields.", vbExclamation
        Exit Sub
    End If

    If IsNumeric(Day.Value) And IsNumeric(Month.Value) And IsNumeric(Year.Value) Then
        dayVal = CInt(Day.Value)
        monthVal = CInt(Month.Value)
        yearVal = CInt(Year.Value)
        
        goalDate = DateSerial(yearVal, monthVal, dayVal)
    Else
        MsgBox "Please enter numeric values for day, month, and year.", vbExclamation
        Exit Sub
    End If
    
    If IsNumeric(Me.InitialAmount.Value) Then
         InitialAmount = CDbl(Me.InitialAmount.Value)
    Else
        MsgBox "input a vaild value"
        Exit Sub
    End If
    
    amountRemaining = InitialAmount
    timeleft = DateDiff("d", Date, goalDate) & " days"

    lastRow = wsGoals.Cells(wsGoals.Rows.Count, 1).End(xlUp).row
    newRowNeeded = True

    'Check if thereas an empty row in the table
    For i = 3 To lastRow
        If WorksheetFunction.CountA(wsGoals.Rows(i)) = 0 Then
            Set firstEmptyRow = wsGoals.Rows(i)
            newRowNeeded = False
            Exit For
        End If
    Next i

    If newRowNeeded Then
        lastRow = lastRow + 1
    Else
        lastRow = firstEmptyRow.row
    End If

    
    wsGoals.Cells(lastRow, 1).Value = Me.GoalName.Value
    wsGoals.Cells(lastRow, 2).Value = goalDate
    wsGoals.Cells(lastRow, 3).Value = timeleft
    wsGoals.Cells(lastRow, 4).Value = InitialAmount
    wsGoals.Cells(lastRow, 5).Value = amountRemaining

    ' Clear form inputs
    Me.GoalName.Value = ""
    Me.Day.Value = ""
    Me.Month.Value = ""
    Me.Year.Value = ""
    Me.InitialAmount.Value = ""

    
    MsgBox "Goal added successfully!", vbInformation
End Sub


Private Sub UserForm_Click()

End Sub
