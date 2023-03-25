Attribute VB_Name = "mTestExpenses"
'@Folder("Expenses")
Option Compare Database
Option Explicit

Public Sub Test()
  Dim e As New CExpenses
  Debug.Print "gros Total: " & e.grosTotal
  Debug.Print "Net Total: " & e.netTotal
  Debug.Print "VAT Total: " & e.vatTotal
End Sub
