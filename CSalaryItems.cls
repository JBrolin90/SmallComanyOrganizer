VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSalaryItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Salary"
Option Compare Database
Option Explicit

Private mEmpId As Long

Private rstItems As DAO.Recordset

Public Property Get EmployeePopularName() As String
  EmployeePopularName = rstItems!PopularName
End Property

Public Property Get CustomerPopularName() As String
  CustomerPopularName = dbReader.CustomerPopularName(rstItems!CustomerID)
End Property

Public Property Get WorkedHours() As String
  WorkedHours = rstItems!DurationPlanned
End Property

Public Property Get Pay() As Double
  Pay = rstItems!Salary
End Property

Public Property Get Price() As Double
  Price = dbReader.HourlyPay(mEmpId)
End Property

Public Property Get ItemDate() As Date
  ItemDate = rstItems!StartTimePlanned
End Property

Public Property Get Count() As Long
  Count = rstItems.RecordCount
End Property

Public Function NextItem() As Boolean
  rstItems.MoveNext
  NextItem = rstItems.EOF
End Function

Public Function RetrieveItems(inEmpID As Long, inYear As Long, inMonth As Long) As Long
  Dim dbs As DAO.Database
  Set dbs = CurrentDb
  Dim qdf As DAO.QueryDef
  
  mEmpId = inEmpID
  Dim mMonth As Long
  mMonth = inMonth
  Dim mYear As Long
  mYear = inYear
  Set qdf = dbs.QueryDefs("SalaryItems")
  qdf.Parameters("Month") = inMonth
  qdf.Parameters("Year") = inYear
  qdf.Parameters("EmployeeID") = inEmpID
  Set rstItems = qdf.OpenRecordset(dbOpenDynaset, dbSeeChanges)
  If Count > 0 Then
    rstItems.MoveFirst
  End If
  rstItems.Sort = "StartTimePlanned"
  RetrieveItems = Count
End Function

