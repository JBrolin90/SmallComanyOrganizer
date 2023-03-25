VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CBookingBrief"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Schedule")
Option Compare Database
Option Explicit

Private Const sqlBase = "SELECT Customers.PopularName AS CPopName, Schedules.StartTimePlanned AS [Date], Schedules.StartTime, Schedules.EndTime, Jobs.JobName, Employees.PopularName AS EPopName, Employees.ID AS EmpID, Customers.ID AS CustID, Schedules.ID AS BookingID FROM Jobs INNER JOIN (Schedules INNER JOIN (Customers INNER JOIN Employees ON Customers.EmployeeID = Employees.EmployeeID) ON (Schedules.EmployeeIDPlanned = Employees.ID) AND (Schedules.CustomerID = Customers.ID)) ON Jobs.ID = Schedules.Job "

Private sql As String

Private Sub Class_Initialize()
  sql = sqlBase
End Sub

Public Property Get rst() As DAO.Recordset
  Set rst = db.GetRecordset(sql)
End Property

Public Function Filter(ByVal CustomerID As String, ByVal employeedID As String) As DAO.Recordset
  sql = sqlBase & " WHERE Customers.ID Like " & CustomerID & " AND Employees.ID like " & employeedID
  Set Filter = rst
End Function

Public Function ClearFilter() As DAO.Recordset
  sql = sqlBase
  Set ClearFilter = rst
End Function
