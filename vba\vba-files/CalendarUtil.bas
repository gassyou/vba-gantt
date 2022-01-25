' Attribute VB_Name = "CalendarUtil"

Option Explicit

Private startDate As Range
Public Duration As Integer
Public YearTitleRow As Integer
Public DateTitleRow As Integer

Public Property Let StartDateRange(val As Range)
    Set startDate = val
    If Not IsDate(val) Or IsEmpty(val) Then
        startDate = Date
    ElseIf DateDiff("d", CDate("2000-01-01"), val) < 0 Then
         startDate = Date
    End If
    
End Property

Public Property Get StartColumnIndex() As Integer
  If IsEmpty(startDate) Then
    StartColumnIndex = 0
  Else
    StartColumnIndex = startDate.Column
  End If
End Property

Public Property Get EndColumnIndex() As Integer
  If IsEmpty(Duration) Then
    EndColumnIndex = 0
  Else
    EndColumnIndex = StartColumnIndex() + Duration
  End If
End Property

Public Sub DrawTitle()

  Dim i As Integer
  For i = 0 To Duration
    'Set Calendar Day
    Dim calendarDay
    calendarDay = DateAdd("d", i, startDate)
    Cells(DateTitleRow, StartColumnIndex() + i) = calendarDay

    'Set Calendar Year-Month
    If day(calendarDay) = 1 Or i = 0 Then
        Cells(YearTitleRow, StartColumnIndex() + i) = calendarDay
    Else
        Cells(YearTitleRow, StartColumnIndex() + i) = ""
    End If
  Next

  With Range(Cells(YearTitleRow, StartColumnIndex()), Cells(YearTitleRow, EndColumnIndex()))
      .ShrinkToFit = True
      .NumberFormat = "yyyy-mm"
  End With
  
  With Range(Cells(DateTitleRow, StartColumnIndex()), Cells(DateTitleRow, EndColumnIndex()))
      .ColumnWidth = 3
      .NumberFormat = "d"
  End With
  
End Sub

Public Sub DrawPlan(ByVal planStartDay As Range, ByVal planEndDay As Range)
  Dim calendarRow As Range
  Set calendarRow = Range(Cells(planStartDay.Row, StartColumnIndex()), Cells(planStartDay.Row, EndColumnIndex()))
  calendarRow.Interior.Pattern = xlNone
  calendarRow.Interior.TintAndShade = 0
  calendarRow.Interior.PatternTintAndShade = 0

  If (IsEmpty(planStartDay) Or IsEmpty(planEndDay) Or Not IsDate(planStartDay) Or Not IsDate(planEndDay)) Then
    Exit Sub
  End If

  If (CDate(planEndDay) < CDate(planStartDay)) Then
    MsgBox ("The plan start day is larger than plan end day!")
    Exit Sub
  End If

  Dim planStartColumn As Integer
  Dim planEndColumn As Integer
  planStartColumn = ColumnIndexInCalendarOfDay(planStartDay)
  planEndColumn = ColumnIndexInCalendarOfDay(planEndDay)

  Dim planRange As Range
  Set planRange = Range(Cells(planStartDay.Row, planStartColumn), Cells(planStartDay.Row, planEndColumn))
  planRange.Interior.Pattern = xlNone
  planRange.Interior.ColorIndex = 23
End Sub

Public Sub DrawActual(ByVal actualStartDay As Range, ByVal actualEndDay As Range, ByVal taskDuration As Range, ByVal completed As Range)
  
  Dim actualBarName As String
  actualBarName = "actual" + CStr(actualStartDay.Row)

  Call ClearRect(actualBarName)

  If (IsEmpty(actualStartDay)) Then
    Exit Sub
  End If

  Dim Bx As Single, By As Single, W As Single, H As Single

  Dim startRange As Range
  Set startRange = RangeInCalendarOfDay(actualStartDay)

  Bx = startRange.left
  H = startRange.RowHeight * 0.5
  W = 0
  By = startRange.top + (H * 0.5)

Dim endRange As Range
  If (Not IsEmpty(actualEndDay) And IsDate(actualEndDay)) Then
    Set endRange = RangeInCalendarOfDay(actualEndDay)
    W = endRange.left + endRange.width - Bx
  ElseIf (Not IsEmpty(taskDuration) And Not IsEmpty(completed)) Then
    Set endRange = startRange.Offset(0, taskDuration.Value - 1)
    W = (endRange.left + endRange.width - Bx) * completed.Value
  End If

  If W = 0 Then
    Exit Sub
  End If

  Call DrawRect(actualBarName, Bx, By, W, H)
End Sub

Public Sub DrawToday()

  Dim calendarArea As Range
  Set calendarArea = Range(Cells(1, StartColumnIndex()), Columns(EndColumnIndex()))

  calendarArea.Font.ColorIndex = xlAutomatic
  
  With calendarArea.Borders(xlEdgeLeft)
    .LineStyle = xlContinuous
    .ThemeColor = 1
    .TintAndShade = -0.14996795556505
    .Weight = xlThin
  End With
  With calendarArea.Borders(xlEdgeRight)
    .LineStyle = xlContinuous
    .ThemeColor = 1
    .TintAndShade = -0.14996795556505
    .Weight = xlThin
  End With
  With calendarArea.Borders(xlInsideVertical)
    .LineStyle = xlContinuous
    .ThemeColor = 1
    .TintAndShade = -0.14996795556505
    .Weight = xlThin
  End With

  Dim columnIndexOfToday As Integer
  columnIndexOfToday = ColumnIndexInCalendarOfDay(Date)

  Dim columnOfToday As Range
  Set columnOfToday = Range(Cells(1, columnIndexOfToday), Range(Cells(1, columnIndexOfToday), Cells(1, columnIndexOfToday)).End(xlDown))

  With columnOfToday.EntireColumn.Borders(xlEdgeLeft)
    .LineStyle = xlDashDot
    .Color = -16776961
    .TintAndShade = 0
    .Weight = xlThin
  End With

  With columnOfToday.EntireColumn.Borders(xlEdgeTop)
    .LineStyle = xlDashDot
    .Color = -16776961
    .TintAndShade = 0
    .Weight = xlThin
  End With

  With columnOfToday.EntireColumn.Borders(xlEdgeBottom)
    .LineStyle = xlDashDot
    .Color = -16776961
    .TintAndShade = 0
    .Weight = xlThin
  End With

  With columnOfToday.EntireColumn.Borders(xlEdgeRight)
    .LineStyle = xlDashDot
    .Color = -16776961
    .TintAndShade = 0
    .Weight = xlThin
  End With
  
  columnOfToday.Font.Color = 255
  
End Sub

Public Function ColumnIndexInCalendarOfDay(ByVal dayRange As Variant)
  Dim intervalDays As Long

  If IsEmpty(startDate) Or IsEmpty(dayRange) Or Not IsDate(startDate) Or Not IsDate(dayRange) Then
    intervalDays = 0
  Else
    intervalDays = DateDiff("d", startDate, dayRange)
  End If

  If intervalDays < 0 Then
    MsgBox ("The Date Need larger than calendar start day! " + dayRange.Address())
    intervalDays = 0
  End If
  ColumnIndexInCalendarOfDay = CLng(StartColumnIndex()) + intervalDays
End Function

Public Function RangeInCalendarOfDay(ByVal dayRange As Range)
  Dim columnIndex As Integer
  columnIndex = ColumnIndexInCalendarOfDay(dayRange)
  Set RangeInCalendarOfDay = Range(Cells(dayRange.Row, columnIndex), Cells(dayRange.Row, columnIndex))
End Function

Private Sub DrawRect(ByVal name As String, ByVal left As Single, ByVal top As Single, ByVal width As Single, ByVal height As Single)
  With ActiveSheet.Shapes.AddShape(msoShapeRectangle, left, top, width, height)
    .name = name
    .Fill.Visible = msoTrue
    .Fill.ForeColor.ObjectThemeColor = msoThemeColorText1
    .Fill.ForeColor.TintAndShade = 0
    .Fill.ForeColor.Brightness = 0
    .Fill.Transparency = 0
    .Fill.Solid
    .Line.Visible = msoTrue
    .Line.ForeColor.ObjectThemeColor = msoThemeColorText1
    .Line.ForeColor.TintAndShade = 0
    .Line.ForeColor.Brightness = 0
    .Line.Transparency = 0
  End With
End Sub

Private Sub ClearRect(ByVal name As String)
  Dim s As Shape
  For Each s In ActiveSheet.Shapes
    If s.name = name Then
      s.Delete
    End If
  Next s
End Sub
