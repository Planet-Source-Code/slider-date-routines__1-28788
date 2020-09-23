Attribute VB_Name = "mDate"
'===========================================================================
'
' Module Name:  mDate
' Author:       Slider
' Date:         11/11/2001
' Version:      01.00.00
' Description:  Date Routines
' Edit History: 01.00.00 11/11/01 Initial Release
'
'===========================================================================

Option Explicit

Private Declare Function VarMonthName Lib "oleaut32" (ByVal lMonth As Long, _
                                                      ByVal fAddrev As Long, _
                                                      ByVal dwFlags As Long, _
                                                      pbstrOut As String) As Long

Private Declare Function VarWeekdayName Lib "oleaut32" (ByVal lWeekDay As Long, _
                                                        ByVal fAddrev As Long, _
                                                        ByVal lFirstDay As Long, _
                                                        ByVal dwFlags As Long, _
                                                        pbstrOut As String) As Long

'===========================================================================
'## Public Date Calculation Routines
'
Public Function FindEOCM(FindDate As Date) As Long
    '## Find last day of current month
    FindEOCM = DateSerial(Year(FindDate), Month(FindDate) + 1, 0)
End Function

Public Function FindEOPM(FindDate As Date) As Long
    '## Find last day of previous month
    FindEOPM = DateSerial(Year(FindDate), Month(FindDate), 0)
End Function

Public Function WeekNumber(InDate As Date) As Integer
    '
    ' Notes: (Microsoft KB Article http://support.microsoft.com/support/kb/articles/Q200/2/99.asp)
    ' ------
    '
    ' ISO 8601 "Data elements and interchange formats - Information interchange   - Representation of dates and times"
    ' ISO 8601 : 1988 (E) paragraph 3.17:
    ' "week, calendar: A seven day period within a calendar year, starting on a Monday and
    ' identified by its ordinal number within the year; the first calendar week of the year
    ' is the one that includes the first Thursday of that year. In the Gregorian calendar,
    ' this is equivalent to the week which includes 4 January."
    '
    ' This can be implemented by applying these rules for Calendar weeks:
    '   - A year is divided into either 52 or 53 calendar weeks.
    '   - A calendar week has 7 days. Monday is day 1, Sunday is day 7.
    '   - The first calendar week of a year is the one containing at least 4 days.
    '   - If a year is not concluded on a Sunday, either its 1-3 last days belong to next
    '     year's first calendar week or the first 1-3 days of next year belong to the
    '     present year's last calendar week.
    '   - Only a year starting or concluding on a Thursday has 53 calendar weeks.

  Dim DayNo     As Integer
  Dim StartDays As Integer
  Dim StopDays  As Integer
  Dim StartDay  As Integer
  Dim StopDay   As Integer
  Dim VNumber   As Integer
  Dim ThurFlag  As Boolean

    DayNo = Days(InDate)
    StartDay = Weekday(DateSerial(Year(InDate), 1, 1)) - 1
    StopDay = Weekday(DateSerial(Year(InDate), 12, 31)) - 1
    ' Number of days belonging to first calendar week
    StartDays = 7 - (StartDay - 1)
    ' Number of days belonging to last calendar week
    StopDays = 7 - (StopDay - 1)
    ' Test to see if the year will have 53 weeks or not
    If StartDay = 4 Or StopDay = 4 Then ThurFlag = True Else ThurFlag = False
    VNumber = (DayNo - StartDays - 4) / 7
    ' If first week has 4 or more days, it will be calendar week 1
    ' If first week has less than 4 days, it will belong to last year's
    ' last calendar week
    If StartDays >= 4 Then
        WeekNumber = Fix(VNumber) + 2
    Else
        WeekNumber = Fix(VNumber) + 1
    End If
    ' Handle years whose last days will belong to coming year's first
    ' calendar week
    If WeekNumber > 52 And ThurFlag = False Then WeekNumber = 1
    ' Handle years whose first days will belong to the last year's
    ' last calendar week
    If WeekNumber = 0 Then
        WeekNumber = WeekNumber(DateSerial(Year(InDate) - 1, 12, 31))
    End If

End Function

Public Function WeekNum2Date(iYear As Integer, iWeekNum As Integer, eWeekDay As DayConstants) As Date

    Dim iStartDay As Integer

    iStartDay = eWeekDay - 1
    If iStartDay = 0 Then iStartDay = 7
    WeekNum2Date = DateSerial(iYear, 1, 7 * (iWeekNum - 1) + iStartDay)

End Function

Public Sub GetMonthNames(ByVal bAbbreviate As Boolean, _
                               MonthNames() As String)

    Dim lLoop       As Long
    Dim fAbbrevFlag As Long
    Dim sResult     As String

    ReDim MonthNames(1 To 12)
    fAbbrevFlag = Abs(bAbbreviate)

    For lLoop = 1 To 12
        sResult = Space$(32)
        VarMonthName lLoop, fAbbrevFlag, 0&, sResult
        MonthNames(lLoop) = StrConv(sResult, vbFromUnicode)
    Next

End Sub

Public Sub GetWeekDayNames(ByVal bAbbreviate As Boolean, _
                                 WeekDayNames() As String, _
                        Optional fFirstDay As DayConstants = ocalMonday)

    Dim lLoop       As Long
    Dim fAbbrevFlag As Long
    Dim lFirstDay   As Long
    Dim sResult     As String

    ReDim WeekDayNames(1 To 7)
    fAbbrevFlag = Abs(bAbbreviate)
    lFirstDay = CLng(fFirstDay)

    For lLoop = 1 To 7
        sResult = Space$(32)
        VarWeekdayName lLoop, fAbbrevFlag, lFirstDay, 0&, sResult
        WeekDayNames(lLoop) = StrConv(sResult, vbFromUnicode)
    Next

End Sub

'===========================================================================
'## Private Date Calculation Routines
'
Private Function Days(DayNo As Date) As Integer
    Days = DayNo - DateSerial(Year(DayNo), 1, 0)
End Function
