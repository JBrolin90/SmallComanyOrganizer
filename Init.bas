Attribute VB_Name = "Init"
'@Folder "HEMS"
Option Compare Database
Option Explicit

Public Property Get AppPath() As String
  AppPath = Application.CurrentProject.Path
End Property

Public Function InitHEMS() As Variant
  SetDB_Access
End Function

Public Sub SetDB_Access()
  Dim dbCurr As DAO.Database
  Dim tdfTableLink As TableDef
  
  Set dbCurr = Application.CurrentDb
  ' https://www.tek-tips.com/viewthread.cfm?qid=1776820
  ' https://stackoverflow.com/questions/20643263/how-can-one-search-tabledefs-for-linked-tables
  ' https://docs.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/tabledefattributeenum-enumeration-dao
  For Each tdfTableLink In dbCurr.TableDefs
    '    If tdfTableLink.Connect <> "" Then
    Debug.Print tdfTableLink.Name & " | " & tdfTableLink.Connect & " | " & tdfTableLink.Attributes
    If tdfTableLink.Attributes = dbAttachedTable Then
      tdfTableLink.Connect = ";DATABASE=" & AppPath & "/HEMS_be.accdb"
      tdfTableLink.RefreshLink
      Debug.Print tdfTableLink.Name & " | " & tdfTableLink.Connect & " | " & tdfTableLink.Attributes
    End If
  Next
End Sub

Public Sub Relink_a_table()
  Dim tbdef As DAO.TableDef
  Dim tdfTableLink As TableDef
  For Each tdfTableLink In CurrentDb.TableDefs
    Debug.Print tdfTableLink.Name & " | " & tdfTableLink.Connect & " | " & tdfTableLink.Attributes
    If tdfTableLink.Name = "Customers" Then
      tdfTableLink.Connect = "ODBC;DRIVER=ODBC Driver 17 for SQL Server;SERVER=192.168.0.139;Trusted_Connection=No;APP=Microsoft Office;DATABASE=HEMS_Econ;UID=sa;PWD=Sqlserver2020"
      tdfTableLink.Attributes = 1073741824
      tdfTableLink.RefreshLink
      Debug.Print tdfTableLink.Name & " | " & tdfTableLink.Connect & " | " & tdfTableLink.Attributes
    End If
  Next

End Sub

