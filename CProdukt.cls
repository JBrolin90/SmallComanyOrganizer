VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CProdukt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Product"
Option Compare Database
Option Explicit

Private rst As DAO.Recordset

Public Sub Product(inProduktID As Long)
  Set rst = db.connection.OpenRecordset("Jobs", dbOpenDynaset, dbSeeChanges)
End Sub

Public Property Get NormalPrice() As Long
  NormalPrice = rst!PriceHigh
End Property

Public Property Get LowPrice() As Long
  LowPrice = rst!PriceLow
End Property

Public Property Get RUT() As Boolean
  RUT = rst!RUT
End Property

