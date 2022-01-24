
Private Sub Worksheet_Change(ByVal target As Range)
    
    CalendarUtil.StartDateRange = Range("M2")
    CalendarUtil.Duration = 3651
    CalendarUtil.YearTitleRow = 1
    CalendarUtil.DateTitleRow = 2
    
    TaskUtil.StartDayOfCalendar = CDate(Range("M2").Value)
    TaskUtil.LinkColumn = 1
    TaskUtil.planStartColumn = 7
    TaskUtil.PlanEndColum = 8
    TaskUtil.ActualStartColumn = 9
    TaskUtil.ActualEndColumn = 10
    TaskUtil.TaskDurationColumn = 11
    TaskUtil.CompletionColumn = 12

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual   'pre XL97 xlManua

    If Not Application.Intersect(target, Range("M2")) Is Nothing Then
        Call CalendarUtil.DrawTitle
    End If
    Call TaskUtil.ScheduleTask(target)
    Call CalendarUtil.DrawPlan(TaskUtil.GetPlanStartRange(target), TaskUtil.GetPlanEndRange(target))
    Call CalendarUtil.DrawActual(TaskUtil.GetActualStartRange(target), TaskUtil.GetActualEndRange(target), TaskUtil.GetTaskDurationRange(target), TaskUtil.GetCompletionRange(target))
    Call CalendarUtil.DrawToday
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic   'pre XL97 xlManua
    
End Sub


