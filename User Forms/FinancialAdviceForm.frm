VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FinancialAdviceForm 
   Caption         =   "UserForm1"
   ClientHeight    =   6285
   ClientLeft      =   120
   ClientTop       =   460
   ClientWidth     =   8900.001
   OleObjectBlob   =   "FinancialAdviceForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FinancialAdviceForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnGoalProgress_Click()
    Call ShowGoalProgress
End Sub

Private Sub btnIncomeExpense_Click()
    Call ShowIncomeExpense
End Sub

Private Sub btnSavingsAnalysis_Click()
    Call ShowSavingsAnalysis
End Sub

Sub ShowGoalProgress()
    Dim wsGoals As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim msg As String
    Dim sumRemaining As Double
    Dim sumInitial As Double
    Dim progress As Double
    Dim goalDate As Date
    Dim daysLeft As Long
    Dim dateValue As String
    Dim remainingValue As Variant

    Set wsGoals = ThisWorkbook.Sheets("Financial Goals")
    lastRow = wsGoals.Cells(wsGoals.Rows.Count, "A").End(xlUp).row

    sumRemaining = 0
    sumInitial = 0
    msg = ""

    For i = 4 To lastRow
        dateValue = wsGoals.Cells(i, 2).Value
        On Error Resume Next
        goalDate = CDate(dateValue)
        On Error GoTo 0

        If Err.Number <> 0 Then
            MsgBox "Error converting date: " & dateValue
            Err.Clear
        Else
            daysLeft = goalDate - Date
            If daysLeft < 7 Then
                msg = msg & wsGoals.Cells(i, 1).Value & " is due in " & daysLeft & " days." & vbCrLf
            End If
            
            ' Check if the amount remaining is numeric
            remainingValue = wsGoals.Cells(i, 5).Value
            If IsNumeric(remainingValue) Then
                sumRemaining = sumRemaining + remainingValue
            Else
                MsgBox "Non-numeric value found in 'Amount Remaining' for goal: " & wsGoals.Cells(i, 1).Value
            End If
            
            ' Check if the initial amount is numeric before summing
            If IsNumeric(wsGoals.Cells(i, 4).Value) Then
                sumInitial = sumInitial + wsGoals.Cells(i, 4).Value
            End If
        End If
    Next i

    If sumInitial <> 0 Then
        progress = (sumRemaining / sumInitial) * 100
    Else
        progress = 0
    End If

    msg = msg & vbCrLf & "Total progress towards goals: " & Format(progress, "0.00") & "%"

    If progress > 50 Then
        msg = msg & vbCrLf & "Great job! You're making good progress towards your goals."
    Else
        msg = msg & vbCrLf & "You might want to save more to meet your goals."
    End If

    MsgBox msg
End Sub
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

Sub ShowSavingsAnalysis()
    Dim wsExpenses As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim categories As Object
    Dim cat As Variant
    Dim maxCategory As String
    Dim maxSpending As Double
    Dim msg As String

    Set wsExpenses = ThisWorkbook.Sheets("Expenses&Incomes")
    lastRow = wsExpenses.Cells(wsExpenses.Rows.Count, "A").End(xlUp).row

    Set categories = CreateObject("Scripting.Dictionary")

    For i = 4 To lastRow
        If wsExpenses.Cells(i, 3).Value <> "Income" Then
            cat = wsExpenses.Cells(i, 3).Value
            If categories.exists(cat) Then
                categories(cat) = categories(cat) + wsExpenses.Cells(i, 4).Value
            Else
                categories.Add cat, wsExpenses.Cells(i, 4).Value
            End If
        End If
    Next i

    maxSpending = 0
    For Each cat In categories.Keys
        If categories(cat) > maxSpending Then
            maxSpending = categories(cat)
            maxCategory = cat
        End If
    Next cat

    msg = "You are spending the most on " & maxCategory & ". Try to cut back on it."

    MsgBox msg
End Sub




Private Sub UserForm_Click()

End Sub
