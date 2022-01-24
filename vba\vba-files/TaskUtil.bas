Option Explicit

Public LinkColumn As Integer
Public planStartColumn As Integer
Public PlanEndColum As Integer
Public ActualStartColumn As Integer
Public ActualEndColumn As Integer
Public TaskDurationColumn As Integer
Public CompletionColumn As Integer
Private CalendarStartDate As Date

Public Property Let StartDayOfCalendar(val As Date)
    CalendarStartDate = val
End Property

Public Property Get StartDayOfCalendar() As Date
    StartDayOfCalendar = CalendarStartDate
End Property

Public Sub ScheduleTask(ByVal editTarget As Range)

    Dim linkRange As Range
    Dim planStartRange As Range, planEndRange  As Range, taskDurationRange As Range
    Dim actualStartRange As Range, actualEndRange As Range, completionRange As Range

    Set linkRange = GetLinkRange(editTarget)

    Set planStartRange = GetPlanStartRange(editTarget)
    Set planEndRange = GetPlanEndRange(editTarget)
    Set taskDurationRange = GetTaskDurationRange(editTarget)

    Set actualStartRange = GetActualStartRange(editTarget)
    Set actualEndRange = GetActualEndRange(editTarget)
    Set completionRange = GetCompletionRange(editTarget)
    
    If (Not IsDate(CalendarStartDate)) Then
        MsgBox ("calendar start day is not correct date")
        Exit Sub
    End If

   Dim intervalDays As Integer
    If editTarget.column = planStartColumn Then
        If (IsEmpty(planStartRange)) Then
            Exit Sub
        End If
        If (Not IsDate(planStartRange)) Then
            MsgBox ("Task plan start day is not correct date")
            Exit Sub
        End If
        If DateDiff("d", CalendarStartDate, planStartRange) < 0 Then
            MsgBox ("Plan start day must be larger than calendar start day!")
            Exit Sub
        End If
        If IsEmpty(planEndRange) And IsEmpty(taskDurationRange) Then
            planEndRange = CalendarStartDate
            taskDurationRange = 1
            Exit Sub
        End If
        If Not IsEmpty(planEndRange) And IsEmpty(taskDurationRange) Then
            intervalDays = DateDiff("d", planStartRange, planEndRange)
            If intervalDays < 0 Then
                MsgBox (" Plan Start day is larger than Plan end day!")
                Exit Sub
            End If
            taskDurationRange = intervalDays + 1
            Exit Sub
        End If
        If Not IsEmpty(taskDurationRange) Then
            planEndRange = DateAdd("d", taskDurationRange - 1, planStartRange)
            Exit Sub
        End If
        
    ElseIf editTarget.column = PlanEndColum Then
        If (IsEmpty(planEndRange)) Then
            Exit Sub
        End If
        If (Not IsDate(planEndRange)) Then
            MsgBox ("Task plan end day is not correct date")
            Exit Sub
        End If
        If DateDiff("d", CalendarStartDate, planEndRange) < 0 Then
            MsgBox ("Plan end day must be larger than calendar start day!")
            Exit Sub
        End If
        If IsEmpty(planStartRange) And IsEmpty(taskDurationRange) Then
            planStartRange = planEndRange
            taskDurationRange = 1
        End If
        If IsEmpty(planStartRange) And Not IsEmpty(taskDurationRange) Then
            planStartRange = DateAdd("d", -(taskDurationRange - 1), planEndRange)
            Exit Sub
        End If
        If Not IsEmpty(planStartRange) Then
            intervalDays = DateDiff("d", planStartRange, planEndRange)
            If intervalDays < 0 Then
                MsgBox (" Plan Start day is larger than Plan end day!")
                Exit Sub
            End If
            taskDurationRange = intervalDays + 1
            Exit Sub
        End If
    ElseIf editTarget.column = TaskDurationColumn Then
        If (IsEmpty(taskDurationRange)) Then
            taskDurationRange = 1
        End If
        If taskDurationRange.Value <= 0 Then
            MsgBox (" Task duration must be larger than 0!")
            Exit Sub
        End If
        If IsEmpty(planStartRange) Then
            planStartRange = CalendarStartDate
        End If
        planEndRange = DateAdd("d", taskDurationRange - 1, planStartRange)
        Exit Sub
    ElseIf editTarget.column = ActualStartColumn Then
        If (IsEmpty(actualStartRange)) Then
            Exit Sub
        End If
        If (Not IsDate(actualStartRange)) Then
            MsgBox ("Task actual start day is not correct date")
            Exit Sub
        End If
        Exit Sub
    ElseIf editTarget.column = ActualEndColumn Then
        If (IsEmpty(actualEndRange)) Then
            Exit Sub
        End If
        If (Not IsDate(actualEndRange)) Then
            MsgBox ("Task actual end day is not correct date")
            Exit Sub
        End If
        If IsEmpty(actualStartRange) Then
            actualStartRange = planStartRange
        End If
        completionRange = 1
        Exit Sub
    ElseIf editTarget.column = CompletionColumn Then
        If (IsEmpty(completionRange)) Then
            Exit Sub
        End If
        If IsEmpty(actualStartRange) Then
            actualStartRange = planStartRange
        End If
        If IsEmpty(actualEndRange) And completionRange = 1 Then
            actualEndRange = planEndRange
        End If
        Exit Sub
    ElseIf editTarget.column = LinkColumn Then
        ' TODOs
    Else
        If IsEmpty(planStartRange) And IsEmpty(planEndRange) Then
            planStartRange = CalendarStartDate
            planEndRange = CalendarStartDate
            taskDurationRange = 1
        End If
    End If
End Sub

Public Function GetLinkRange(ByVal target As Range) As Range
    Set GetLinkRange = Range(Cells(target.Row, LinkColumn), Cells(target.Row, LinkColumn))
End Function

Public Function GetPlanStartRange(ByVal target As Range) As Range
    Set GetPlanStartRange = Range(Cells(target.Row, planStartColumn), Cells(target.Row, planStartColumn))
End Function

Public Function GetPlanEndRange(ByVal target As Range) As Range
    Set GetPlanEndRange = Range(Cells(target.Row, PlanEndColum), Cells(target.Row, PlanEndColum))
End Function

Public Function GetTaskDurationRange(ByVal target As Range) As Range
    Set GetTaskDurationRange = Range(Cells(target.Row, TaskDurationColumn), Cells(target.Row, TaskDurationColumn))
End Function

Public Function GetActualStartRange(ByVal target As Range) As Range
    Set GetActualStartRange = Range(Cells(target.Row, ActualStartColumn), Cells(target.Row, ActualStartColumn))
End Function

Public Function GetActualEndRange(ByVal target As Range) As Range
    Set GetActualEndRange = Range(Cells(target.Row, ActualEndColumn), Cells(target.Row, ActualEndColumn))
End Function

Public Function GetCompletionRange(ByVal target As Range) As Range
    Set GetCompletionRange = Range(Cells(target.Row, CompletionColumn), Cells(target.Row, CompletionColumn))
End Function



