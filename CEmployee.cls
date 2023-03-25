VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Employee")
Option Compare Database
Option Explicit

Private Const ErrBase = 600
Private mEmpId As String

Private dbs As DAO.Database
Private rstEmployee As DAO.Recordset

Public Sub EmpID(Optional ByVal inEmpID As String = "*")
  mEmpId = inEmpID
  Set dbs = CurrentDb
  RetrieveEmployee
End Sub

Public Function getEmpID() As Long
  getEmpID = rstEmployee!ID
End Function

Public Property Get Count() As Long
    Count = rstEmployee.RecordCount
End Property

Public Sub Begin()
    rstEmployee.MoveFirst
End Sub

Public Function nextEmployee() As Boolean
  nextEmployee = True
  rstEmployee.MoveNext
  If rstEmployee.EOF Then nextEmployee = False
End Function


Public Property Get FirstName() As String
  FirstName = rstEmployee!FirstName
End Property

Public Property Get LastName() As String
  LastName = rstEmployee!LastName
End Property

Public Property Get FullName() As String
  FullName = FirstName & " " & LastName
End Property

Public Property Get PopularName() As String
  PopularName = rstEmployee!PopularName
End Property


Public Property Get TaxTable() As Long
  TaxTable = rstEmployee!TaxTable
End Property

Public Property Get SocialFeeRate() As Double
  SocialFeeRate = rstEmployee!SocialFee
End Property

Public Property Get PayDay() As Long
On Error GoTo err
  PayDay = rstEmployee!PayDay
Exit Property
err:
  Debug.Print "ERROR: " & err.Number & " Description: " & err.Description & " source: " & err.Source
End Property

Public Property Get Address() As String
On Error GoTo err
  Address = dbReader.EmpAddress(getEmpID())
Exit Property
err:
  Raise ErrBase + 3, err.Number & " - " + err.Description + err.Source, "Address", "CEmployee"
End Property

Public Property Get StatementFolder() As String
On Error GoTo err
  If IsNull(rstEmployee!StatementFolder) Then
    DefaultStatementFolder
  End If
  StatementFolder = Settings.StatmentBaseFolder + rstEmployee!StatementFolder
Exit Property
err:
  Debug.Print "ERROR: " & err.Number & " Description: " & err.Description & " source: " & err.Source
  Raise ErrBase + 1, err.Number & " - " + err.Description + err.Source, "StatementFolder", "CEmployee"
End Property

Public Property Get SchemaFolder() As String
On Error GoTo err
  If IsNull(rstEmployee!StatementFolder) Then
    DefaultStatementFolder
  End If
  SchemaFolder = Settings.SchemaBaseFolder & rstEmployee!StatementFolder & "..\Schema\"
Exit Property
err:
  Debug.Print "ERROR: " & err.Number & " Description: " & err.Description & " source: " & err.Source
  Raise ErrBase + 1, err.Number & " - " + err.Description + err.Source, "StatementFolder", "CEmployee"
End Property

Private Sub DefaultStatementFolder()
On Error GoTo err
  rstEmployee.Edit
  rstEmployee!StatementFolder = FullName + "\Lönespecifikation\"
  rstEmployee.Update
Exit Sub
err:
  Raise ErrBase + 2, err.Number & " - " + err.Description + err.Source, "DefaultStatementFolder", "CEmployee"
End Sub

Public Sub Init()
  Set dbs = CurrentDb
  RetrieveEmployee
End Sub

Private Sub RetrieveEmployee()
  Dim qdf As DAO.QueryDef
  Set qdf = dbs.QueryDefs.Item("Employee")
  qdf.Parameters("EmpID") = mEmpId
  Set rstEmployee = qdf.OpenRecordset(dbOpenDynaset, dbSeeChanges)
End Sub

