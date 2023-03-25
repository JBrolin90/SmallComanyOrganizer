Attribute VB_Name = "mTestSugeySalary"
'@Folder("SugeySalary")
Option Compare Database
Option Explicit

Public Sub Test()
  Dim ss As New CSugeySalary
  ss.Init 2022, 3
  Debug.Print "GrossSalary: " & ss.GrossSalary
  Debug.Print "Tax        : " & ss.Tax
  Debug.Print "NetSalary: " & ss.NetSalary
End Sub

Public Sub testForm()
  mEventProxy.EventProxy.SendPeriod 2022, 3
End Sub
