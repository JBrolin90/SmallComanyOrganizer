VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CControlsConsistency"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Utilities")
Option Compare Database
Option Explicit


Public Function VerifyConsistency(controls As controls) As Boolean
  Dim c As Control
  VerifyConsistency = True
  For Each c In controls
    VerifyConsistency = VerifyConsistency And ConsistencyDispatch(c)
  Next
End Function


Private Function checkTextBox(tb As Access.TextBox) As Boolean
  If IsNull(tb.value) Then
    tb.BorderColor = vbRed
    checkTextBox = False
  Else
    tb.BorderColor = vbBlack
    checkTextBox = True
  End If
End Function

Private Function checkComboBox(cb As Access.ComboBox) As Boolean
   If IsNull(cb.value) Then
    cb.BorderColor = vbRed
    checkComboBox = False
  Else
    cb.BorderColor = vbBlack
    checkComboBox = True
  End If
End Function


Private Function ConsistencyDispatch(c As Control) As Boolean
  ConsistencyDispatch = True
  If c.ControlType = acTextBox Then ConsistencyDispatch = checkTextBox(c)
  If c.ControlType = acComboBox Then ConsistencyDispatch = checkComboBox(c)
End Function
