VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CustomersToInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Invoice")
Option Compare Database
Option Explicit
Private rstItems As DAO.Recordset

Public Function sqlListCustomers(inYear As Integer, inMonth As Integer) As String
  Dim sql As String
  sql = "Select DISTINCT Schedules.CustomerId, PopularName From Schedules "
  sql = sql + "Inner Join Customers ON Schedules.CustomerID = Customers.Id "
  sql = sql + "Where Year(StartTimePlanned)=" & inYear & " "
  sql = sql + "AND Month(StartTimePlanned)=" & inMonth & " "
  sqlListCustomers = sql
End Function

Public Function ListOfCustomersToInvoice(inYear As Integer, inMonth As Integer) As DAO.Recordset
  Dim sql As String
  sql = sqlListCustomers(inYear, inMonth)

  Set rstItems = db.connection.OpenRecordset(sql, dbOpenDynaset, dbSeeChanges)
  If rstItems.RecordCount > 0 Then
    rstItems.MoveLast
    rstItems.MoveFirst
  End If
  
  Set ListOfCustomersToInvoice = rstItems
End Function


