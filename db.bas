Attribute VB_Name = "db"
'@Folder("Utilities")
Option Compare Database
Option Explicit
Private dBase As DAO.Database

Public Function OpenDB() As DAO.Database
  Set dBase = DAO.OpenDatabase(vbNullString, False, False, "Driver={ODBC Driver 17 for SQL Server};Server=192.168.0.139;Database=HEMS_Econ;UID=sa;PWD=Sqlserver2020;")
  Set OpenDB = dBase
End Function

Public Property Get connection() As DAO.Database
  If dBase Is Nothing Then OpenDB
  Set connection = dBase
End Property

Public Function GetRecordset(inSQL As String) As DAO.Recordset
  Dim d As DAO.Database
  Dim r As DAO.Recordset
  Set d = connection
  Set r = d.OpenRecordset(inSQL, dbOpenDynaset, dbSeeChanges)
  Set GetRecordset = r
End Function
