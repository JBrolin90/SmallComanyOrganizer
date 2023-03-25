Attribute VB_Name = "mDateFunctions"
'@Folder "Utilities"
Option Compare Database
Option Explicit

Public Function GetLastDayOfMonth(ByVal myDate As Date) As Date
  GetLastDayOfMonth = LastDay(myDate)
End Function

Public Function DaysOfMonth(ByVal inYear As Long, ByVal inMonth As Long) As Long
  DaysOfMonth = Day(LastDay(DateSerial(inYear, inMonth, 1)))
End Function

Public Function LastDay(ByVal d As Date) As Date

  Dim returnDate As Date
  
  'First day of current month
  returnDate = DateSerial(Year(d), Month(d), 1)
  'Forward a month
  returnDate = DateAdd("m", 1, returnDate)
  'back one day
  returnDate = DateAdd("d", -1, returnDate)
  LastDay = returnDate
End Function

Public Function WeekDayStr(inYear As Long, inMonth As Long, inDay As Long) As String
  Dim dte As Date
  Dim wDay As Long
  dte = DateSerial(inYear, inMonth, inDay)
  wDay = Weekday(dte, FirstDayofWeek:=vbMonday)
  WeekDayStr = WeekDayName(wDay, Abbreviate:=True, FirstDayofWeek:=vbMonday)
End Function
Public Function WeekDayNr(inYear As Long, inMonth As Long, inDay As Long) As Long
  Dim dte As Date
  Dim wDay As Long
  dte = DateSerial(inYear, inMonth, inDay)
  wDay = Weekday(dte, FirstDayofWeek:=vbMonday)
  WeekDayNr = wDay
End Function

Public Function WeekNr(ByVal inDate As Date) As Long
  WeekNr = Format(inDate, "ww", vbMonday, vbFirstFourDays)
End Function

Public Function FirstDateOfWeek(inWeekNr As Long) As Date
  Debug.Assert False
End Function

Private Function WNtoDate(WN As Long, YR As Long, Optional DOW As Long = 2) As Date
  'DOW:  1=SUN, 2=MON, etc
  Dim DY1 As Date
  Dim Wk1DT1 As Date
  Dim i As Long

  DY1 = DateSerial(YR, 1, 1)
  'Use ISO weeknumber system
  i = DatePart("ww", DY1, vbMonday, vbFirstFourDays)

  'Sunday of Week 1
  Wk1DT1 = DateAdd("d", -Weekday(DY1), DY1 + 1 + IIf(i > 1, 7, 0))

  WNtoDate = DateAdd("ww", WN - 1, Wk1DT1) + DOW - 1

End Function

