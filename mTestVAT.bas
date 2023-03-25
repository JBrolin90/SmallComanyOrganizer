Attribute VB_Name = "mTestVAT"
'@Folder("Monthly")
Option Compare Database
Option Explicit

Public Sub testVATMonth()
  Dim vm As New CVAT
  Debug.Print vm.VATMonth(2022, 5)
End Sub
