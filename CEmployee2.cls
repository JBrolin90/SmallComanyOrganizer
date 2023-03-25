VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CEmployee2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Employee")
Option Compare Database
Option Explicit

Private mFirstname As String
Private mLastname As String
Private mEmployeeID As Long

Public Property Get EmployeeID() As Long
  EmployeeID = mEmployeeID
End Property
Public Property Let EmployeeID(eID As Long)
  mEmployeeID = eID
End Property

Public Property Get FirstName() As String
  FirstName = mFirstname
End Property
Public Property Let FirstName(fName As String)
  mFirstname = fName
End Property

Public Property Get LastName() As String
  LastName = mLastname
End Property
Public Property Let LastName(lName As String)
  mLastname = lName
End Property



Public Function Employee2rst(Employee As DAO.Fields) As DAO.Fields
  With Employee
    !EmployeeID = EmployeeID
    !FirstName = FirstName
    !LastName = LastName
  End With
End Function

