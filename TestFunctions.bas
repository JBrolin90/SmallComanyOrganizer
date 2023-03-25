Attribute VB_Name = "TestFunctions"
'@Folder("TestFunctions")
Option Compare Database
Option Explicit

Public Sub openExcel()
  Dim e As New CExcel
  e.openExcel "C:\Users\EMWJOAR\OneDrive - Ericsson\Documents\Private\HEMS\2022-01-25_Lonespec_december Jorge Leiva.xlsx"
End Sub

Public Sub CreateSM()
  Dim ss As New CSalaryStatement
  ss.CreateSalaryStatement 3, 2022, 1
End Sub

Public Sub ConnectionStudy()
  Dim db As DAO.Database
  Set db = DAO.OpenDatabase(vbNullString, False, False, "Driver={ODBC Driver 17 for SQL Server};Server=Book.brolin.org;Database=HEMS_Econ;UID=sa;PWD=Sqlserver2020;")
End Sub

Public Sub testCprod()
  Dim p As New CProdukt
  p.Product 1
End Sub

Public Sub TestInvoice()
  Dim inv As New CInvoice
  inv.CreateInvoice 7, 2022, 2
End Sub


Public Sub testWeekNr()
  Dim d As Date
  Debug.Print Format$(DateSerial(2022, 2, 25), "ww", vbMonday, vbFirstFourDays);
End Sub

Public Function Modulename() As String
  Modulename = Application.VBE.ActiveCodePane.CodeModule.Name
End Function


Public Sub TestSchema()
  Dim sch As New CSchemaExcel
  sch.CreateSchema 2, 2022, 1
  sch.SaveSchema
  sch.CloseMe
End Sub
