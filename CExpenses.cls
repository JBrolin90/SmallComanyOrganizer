VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CExpenses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Expenses")
Option Compare Database
Option Explicit

Private rst As DAO.Recordset

Private Sub Init()
  Dim sql As String
  sql = "SELECT * FROM Expenses"
  Set rst = db.GetRecordset(sql)
End Sub

Public Property Get vatTotal() As Double
  rst.MoveFirst
  vatTotal = 0
  Do
    vatTotal = vatTotal + rst!Amount * (1 - (1 / (1 + rst!VAT)))
    'Debug.Print "Amount: " & rst!Amount & " VAT: " & rst!VAT & " vatTotal: " & vatTotal
  rst.MoveNext
  Loop Until rst.EOF
End Property

Public Property Get netTotal() As Double
  rst.MoveFirst
  netTotal = 0
  Do
    netTotal = netTotal + rst!Amount * (1 / (1 + rst!VAT))
    'Debug.Print "Amount: " & rst!Amount & " VAT: " & rst!VAT & " netTotal: " & netTotal
  rst.MoveNext
  Loop Until rst.EOF
End Property

Public Property Get grosTotal() As Double
  rst.MoveFirst
  grosTotal = 0
  Do
    grosTotal = grosTotal + rst!Amount
    'Debug.Print "Amount: " & rst!Amount & " VAT: " & rst!VAT & " grosTotal: " & grosTotal
  rst.MoveNext
  Loop Until rst.EOF
End Property

Private Sub Class_Initialize()
  Init
End Sub
