VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Schedule"
Option Compare Database
Option Explicit
Private rstSch As DAO.Recordset
Private mCustomer As New CCustomer

Public Sub Init(inEmpID As Long, inYear As Long, inMonth As Long)
  Dim s As String
  
  s = "SELECT * FROM Schedules "
  s = s & "WHERE EmployeeIDPlanned = " & inEmpID & " "
  s = s & "AND Year(StartTimePlanned) = " & inYear & " "
  s = s & "AND Month(StartTimePlanned) = " & inMonth & " "
  s = s & "ORDER BY StartTimePlanned ASC"
  Debug.Print s
  Set rstSch = db.connection.OpenRecordset(s, dbOpenDynaset, dbSeeChanges)
  
  
End Sub

Private Sub PopulateCustomer()
  mCustomer.CustomerID rstSch!CustomerID
End Sub


Public Property Get Initials() As String
  PopulateCustomer
  Initials = mCustomer.Initials
End Property

Public Property Get Duration() As Double
  Duration = rstSch!DurationPlanned
End Property

Public Property Get CutomerFullname() As String
  PopulateCustomer
  With mCustomer
    CutomerFullname = .Initials & "-" & .FullName & ", " & .FullAddress
  End With
End Property

Public Property Get DayNr() As Long
  DayNr = Day(rstSch!StartTimePlanned)
End Property

Public Function StepNext() As Boolean
  rstSch.MoveNext
  StepNext = Not rstSch.EOF
End Function

Public Function Beginning() As Boolean
  If rstSch.RecordCount > 0 Then
    rstSch.MoveFirst
    Beginning = True
  Else
    Beginning = False
  End If
End Function


