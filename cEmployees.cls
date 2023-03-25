VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cEmployees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Employee")
Option Compare Database
Option Explicit

Private Const sql As String = "SELECT * FROM Employees"
Private rst As DAO.Recordset
Dim fs As Fields

Private Property Get employees() As DAO.Recordset
  If rst Is Nothing Then
    Set rst = db.GetRecordset(sql)
  End If
  Set employees = rst
End Property


Public Sub Add(ByRef Employee As CEmployee2)
  employees.AddNew
  Employee.Employee2rst employees.Fields
  employees.Update
End Sub

