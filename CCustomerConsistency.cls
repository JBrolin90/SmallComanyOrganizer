VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCustomerConsistency"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Customer")
Option Compare Database
Option Explicit

Public Function frmConsistent(f As Form) As Boolean
  Dim consistency As New CControlsConsistency
  Dim cs As controls
  Set cs = f.controls
  frmConsistent = consistency.VerifyConsistency(cs)
End Function

Public Function Consistent(record As DAO.Fields) As Boolean
  err.Raise 0, "CCustomerConsistency.Consistent", "Not yet implemented"
End Function



Public Function ConfirmSave(ByRef f As Access.Form) As Boolean
  ConfirmSave = True
  Dim reply As Long
  If f.Dirty Then
    reply = MsgBox("Vill du spara dina ändringar?", vbYesNo, "Konfirmera")
    If reply = vbYes Then
      ConfirmSave = True
    Else
      ConfirmSave = False
    End If
  End If
End Function
