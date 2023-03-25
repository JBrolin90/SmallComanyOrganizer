VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSchema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Schema"
Option Compare Database
Option Explicit

'Create schema, data from database, for one employee, to output in, for example an excel file
'It is a table over one month two rows per customer, one row for day of month, one row for weekdays and one row for week numbers
'One textbox with customer information and a place for the employee's name
'The template contains Sincera's logotype in the upper right corner


Private rstEmployee As DAO.Recordset


Private Sub Init(inYear As Long, inMonth As Long, Optional inEmpID As String = "*")
  Dim s As String
  s = "SELECT * FROM Schedules WHERE"
  s = s & "EmpID Like " & inEmpID & " AND "
  s = s & "Year(StartTimePlanned) = " & inYear & " AND "
  s = s & "Month(StartTimePlanned) = " & inMonth & " "
  Set rstEmployee = db.connection.OpenRecordset(s, dbOpenDynaset, dbSeeChanges)
End Sub

