Attribute VB_Name = "mCustomersToInvoice"
'@Folder("Invoice")
Option Compare Database
Option Explicit

Public Sub Test(inYear As Integer, inMonth As Integer)
  Dim customers As CustomersToInvoice
  Dim rst As DAO.Recordset
  Set customers = New CustomersToInvoice

  Set rst = customers.ListOfCustomersToInvoice(inYear, inMonth)
  Dim i As Integer
  Debug.Print
  For i = 1 To rst.RecordCount
    Debug.Print rst!CustomerID & " " & rst!PopularName
    rst.MoveNext
  Next i
  Debug.Print rst.RecordCount & " Customer(s) has invoices in " & inYear & "-" & inMonth

End Sub
