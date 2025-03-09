Attribute VB_Name = "Module14"

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
    Dim daysremaining As String
    Dim j As Long
    Dim PercentLeft As Double
    Dim GoalProjection As Double
    Dim InitialAmount As Double

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
            If daysLeft < 7 And daysLeft > 0 Then
                msg = msg & wsGoals.Cells(i, 1).Value & " is due in " & daysLeft & " days." & vbCrLf
            End If
            
            ' Check if the amount remaining is numeric
            remainingValue = wsGoals.Cells(i, 6).Value
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
    
    For j = 4 To lastRow
        PercentLeft = wsGoals.Cells(j, 7).Value
        InitialAmount = wsGoals.Cells(j, 4).Value
        dateValue = wsGoals.Cells(j, 2).Value
        On Error Resume Next
        goalDate = CDate(dateValue)
        On Error GoTo 0
        
        daysLeft = goalDate - Date
        GoalProjection = PercentLeft * 100 * daysLeft
        GoalReccomendation = (InitialAmount / 2 - GoalProjection) / InitialAmount * 100
        If GoalProjection < InitialAmount / 2 And daysLeft > 0 Then
            msg = msg & wsGoals.Cells(j, 1).Value & " is projected to not be finished by desired date; invest " & Format(GoalReccomendation, "0.00") & "% more into said goal" & vbCrLf
        End If
    Next j
    
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
