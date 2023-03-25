VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Customer"
Option Compare Database
Option Explicit

Private rst As DAO.Recordset

Public Function CustomerID(Optional ByVal inCustomerID As Long = 0) As Boolean
  Dim ID As String
  ID = inCustomerID
  If inCustomerID = 0 Then ID = "'*'"
  Set rst = db.GetRecordset("SELECT * FROM Customers where ID like " & ID)
  If Not rst.EOF Then
    rst.MoveLast
    rst.MoveFirst
    CustomerID = True
  Else
    CustomerID = False
  End If
End Function

Public Property Get Count() As Long
    Count = rst.RecordCount
End Property

Public Sub Begin()
    rst.MoveFirst
End Sub

Public Function nextCustomer() As Boolean
    nextCustomer = True
    rst.MoveNext
    If rst.EOF Then nextCustomer = False
End Function

Public Function getMaxCustomerID() As Long
  Dim rst As DAO.Recordset
  Set rst = db.GetRecordset("SELECT MAX([CustomerID]) as maxid FROM Customers")
  getMaxCustomerID = rst!maxid
End Function
Public Function getCustomerID() As Long
  getCustomerID = rst.Fields().Item("ID").value
End Function

Public Property Get InvoiceFolder() As String
  InvoiceFolder = Settings.InvoiceBaseFolder & rst!InvoiceFolder
End Property

Public Property Get ProductID() As Long
  ProductID = rst!ProduktId
End Property
Public Property Get RUT() As Boolean
  RUT = rst!RUT
End Property

Public Property Get DayOfWeek() As Long
  DayOfWeek = rst!DayOfWeek
End Property
Public Property Get NormalPrice() As Currency
  NormalPrice = rst!PriceHigh
End Property
Public Property Get LowPrice() As Currency
  LowPrice = rst!PriceLow
End Property
Public Property Get EmployeeID() As Long
  EmployeeID = rst!EmployeeID
End Property

Public Property Get FirstName() As String
  FirstName = rst!FirstName
End Property
Public Property Get LastName() As String
  LastName = rst!LastName
End Property
Public Property Get FullName() As String
  FullName = FirstName & " " & LastName
End Property
Public Property Get PopularName() As String
  PopularName = rst!PopularName
End Property

Public Property Get Initials() As String
  Initials = Left$(FirstName, 1) & Left$(LastName, 2)
End Property
Public Property Get PersonalNumber() As String
  PersonalNumber = rst!PersonalNumber
End Property

Public Property Get Street() As String
  Street = rst!Street
End Property
Public Property Get StreetNumber() As String
  StreetNumber = rst!StreetNumber
End Property
Public Property Get Address() As String
  Address = Street & " " & StreetNumber
End Property
Public Property Get FullAddress() As String
  FullAddress = Address & ", " & Zip & " " & City
End Property

Public Property Get Zip() As String
  Zip = rst!Zip
End Property
Public Property Get City() As String
  City = rst!City
End Property

