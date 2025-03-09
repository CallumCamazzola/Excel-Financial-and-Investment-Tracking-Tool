VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GoalContributions 
   Caption         =   "UserForm1"
   ClientHeight    =   6285
   ClientLeft      =   120
   ClientTop       =   460
   ClientWidth     =   8900.001
   OleObjectBlob   =   "GoalContributionsForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GoalContributions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub UserForm_Initialize()
    Dim wsGoals As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    
    Set wsGoals = ThisWorkbook.Sheets("Financial Goals")
    
    ' Find the last row with a goal
    lastRow = wsGoals.Cells(wsGoals.Rows.Count, 1).End(xlUp).row
    
    
    CmbCategory.Clear
    
   
    For i = 4 To lastRow
        If wsGoals.Cells(i, 7).Value <> 1 Then
            CmbCategory.AddItem wsGoals.Cells(i, 1).Value
        End If
    Next i
End Sub

Private Sub CommandButton1_Click()
    Dim wsGoals As Worksheet
    Dim lastRow As Long
    Dim goalRow As Long
    Dim contribution As Double
    Dim selectedGoal As String
    Set wsGoals = ThisWorkbook.Sheets("Financial Goals")
    Set wsBalance = ThisWorkbook.Sheets("Expenses&Incomes")
    
    If CmbCategory.Value = "" Then
        MsgBox " please select a goal"
        Exit Sub
    End If
    selectedGoal = Me.CmbCategory.Value
    If IsNumeric(Me.txtItem.Value) Then
        contribution = CDbl(Me.txtItem.Value)
    Else
        MsgBox "Please Enter a Valid Contribution"
        Exit Sub
    End If
    
    ' Find the row with the selected goal
    lastRow = wsGoals.Cells(wsGoals.Rows.Count, 1).End(xlUp).row
    goalRow = 0
    For i = 2 To lastRow
        If wsGoals.Cells(i, 1).Value = selectedGoal Then
            goalRow = i
            Exit For
        End If
    Next i
    
    If goalRow = 0 Then
        MsgBox "Goal not found. Please select a valid goal.", vbExclamation
        Exit Sub
    End If
    
    If contribution > wsGoals.Cells(goalRow, 5).Value Then
        contribution = wsGoals.Cells(goalRow, 5).Value
    End If
    
    ' Apply the contribution
    wsGoals.Cells(goalRow, 5).Value = wsGoals.Cells(goalRow, 5).Value - contribution
    currentBalance = wsGoals.Range("K2").Value
    wsGoals.Range("K2").Value = currentBalance + contribution
    wsGoals.Cells(goalRow, 6) = wsGoals.Cells(goalRow, 4).Value - wsGoals.Cells(goalRow, 5).Value
    wsGoals.Cells(goalRow, 7) = wsGoals.Cells(goalRow, 6).Value / wsGoals.Cells(goalRow, 4).Value
    MsgBox "Contribution added to " & selectedGoal & " successfully!", vbInformation
    CmbCategory.Value = ""
    txtItem.Value = ""
End Sub
