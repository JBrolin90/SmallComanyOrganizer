VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSchemaExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Schema"
Option Compare Database
Option Explicit

Private mSchemaYear As Long
Private mSchemaMonth As Long
Private mExcel As CExcel
Private mEmployee As New CEmployee
Private mSchedule As New CSchedule

Public Sub CreateSchema(inEmpID As Long, inSchemaYear As Long, inSchemaMonth As Long)
  mEmployee.EmpID inEmpID
  mEmployee.Init
  
  
  mSchemaMonth = inSchemaMonth
  mSchemaYear = inSchemaYear
  
  Dim mTemplate As String
  mTemplate = Settings.SchemaTemplate
  Set mExcel = New CExcel
  mExcel.openExcel mTemplate
  
  PrepareMonth
  
  mSchedule.Init inEmpID, inSchemaYear, inSchemaMonth
  EmitSchema
  EmitCustomers
 
End Sub

Private Sub ClearSchema()
  Dim i As Long
  For i = 2 To 35
    mExcel.Cell(7, i) = vbNullString
    mExcel.Cell(8, i) = vbNullString
  Next
End Sub

Private Sub EmitCustomers()
  Dim i As Long
  Dim j As Long
  Dim NewCustomer As Boolean
  If mSchedule.Beginning Then
    Do
      NewCustomer = True
      For j = 14 To 35
        If mExcel.Cell(j, 14) = mSchedule.CutomerFullname Then NewCustomer = False
      Next
      DoEvents
      Debug.Print mSchedule.CutomerFullname
      If NewCustomer Then
        mExcel.Cell(14 + i, 14) = mSchedule.CutomerFullname
        i = i + 1
      End If
    Loop While mSchedule.StepNext
  End If
End Sub

Private Function FindFreeRow(ByVal inRow As Long, ByVal inCol As Long) As Long
  While mExcel.Cell(inRow, inCol) <> vbNullString
    inRow = inRow + 2
  Wend
  FindFreeRow = inRow
End Function

Private Sub EmitSchema()
  ClearSchema
  With mSchedule
    If .Beginning Then
      Do
        Dim row As Long
        row = FindFreeRow(7, .DayNr + 1)
        mExcel.Cell(row, .DayNr + 1) = .Initials
        mExcel.Cell(row + 1, .DayNr + 1) = .Duration
      Loop While mSchedule.StepNext
    End If
  End With
End Sub

Private Sub PrepareMonth()
  Dim e As New CEmployee
  AddDays
  mExcel.Cell(3, 3) = Format$(DateSerial(mSchemaYear, mSchemaMonth, 1), "mmm-yyyy")
  mExcel.Cell(14, 5) = mEmployee.FullName
End Sub


Private Sub AddDays()
  Dim days As Long
  Dim i As Long
  days = DaysOfMonth(mSchemaYear, mSchemaMonth)
  Debug.Assert days > 27 And days < 32
  For i = 1 To days
    mExcel.Cell(6, i + 1) = i
    mExcel.Cell(5, i + 1) = mDateFunctions.WeekDayStr(mSchemaYear, mSchemaMonth, i)
    'mExcel.Border(4, i + 1) = xlNone
    If WeekDayNr(mSchemaYear, mSchemaMonth, i) = 7 Then
      mExcel.Border(4, i + 1, xlEdgeRight) = xlContinuous
    Else
      mExcel.Border(4, i + 1, xlEdgeRight) = xlNone
    End If
    mExcel.Cell(4, i + 1) = vbNullString 'ClearWeekNr
    
    DoEvents
  Next
  AddWeekNumbers days
End Sub

Private Sub AddWeekNumbers(inDays As Long)
  Dim i As Long
  Dim d As Long
  Dim wr As Long
  Dim wk As Long
  wk = mDateFunctions.WeekNr(DateSerial(mSchemaYear, mSchemaMonth, 1))
  For i = 1 To inDays
    wr = WeekDayNr(mSchemaYear, mSchemaMonth, i)
    wk = mDateFunctions.WeekNr(DateSerial(mSchemaYear, mSchemaMonth, i))
    If wr = 7 Then
      d = (i Mod 7) / 2
      If i >= 7 Then d = 3
      If d < 1 Then d = 1
      mExcel.Cell(4, i + 1 - d) = "V" & wk
      wk = wk + 1
    ElseIf i >= inDays Then
      If mDateFunctions.WeekNr(DateSerial(mSchemaYear, mSchemaMonth, i)) > wk Then
        mExcel.Cell(4, i + 1 - (wr / 2)) = "V" & wk
      End If
    End If
  Next
End Sub

Private Function BuildFileName() As String
  Dim s As String
  Dim d As String
  d = DateSerial(mSchemaYear, mSchemaMonth, 1)
  s = mEmployee.SchemaFolder & mSchemaYear & "\"
  s = s & mSchemaYear & "-" & Format$(mSchemaMonth, "00")
  BuildFileName = s & " " & mEmployee.FullName & "-Schema"
End Function



Public Sub Save()
  SaveSchema
End Sub

Public Sub CloseMe()
  mExcel.CloseWB
End Sub

Public Sub SaveSchema()
  mExcel.SaveAs BuildFileName()
End Sub
