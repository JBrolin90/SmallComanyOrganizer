Attribute VB_Name = "mErr"
'@Folder "Utilities"
Option Compare Database
Option Explicit

Public Const cCaseErrBase As Long = 200
Private Const sql = "SELECT * FROM Error"
Private rst As DAO.Recordset
Private mInstance As Long

Private Property Get errors() As DAO.Recordset
  If rst Is Nothing Then
    Set rst = db.GetRecordset(sql)
  End If
  Set errors = rst
End Property

Private Property Get instance() As Long
  If mInstance = 0 Then
    Dim r As DAO.Recordset
    Set r = db.GetRecordset("SELECT MAX(eInstance) as MaxInstance FROM Error")
    mInstance = Nz(r!MaxInstance, 0) + 1
  End If
  instance = mInstance
End Property

Public Sub LogDebug(errProcedure As String, errModule As String, inDescription As String)
  errors.AddNew
  errors!eInstance = -1
  errors!eDate = Now
  'errors!eUser = GetUsername
  errors!eNumber = 0
  errors!eSource = "mError"
  errors!eLine = Erl
  errors!eDescription = inDescription
  errors!eProcedure = errProcedure
  errors!eModule = errModule
  errors.Update
End Sub

Public Sub LogError(errProcedure As String, errModule As String)
  errors.AddNew
  errors!eInstance = -2
  errors!eNumber = err.Number
  errors!eSource = err.Source
  'errors!eUser = GetUsername
  errors!eLine = Erl
  errors!eDescription = err.Description
  errors!eProcedure = errProcedure
  errors!eModule = errModule
  errors.Update
End Sub

Public Sub LogRaise(errProcedure As String, errModule As String)
  LogError errProcedure, errModule
  err.Raise err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext
End Sub

Public Sub Raise(errNumber As Integer, errDescription As String, errProcedure As String, errModule As String)
  errors.AddNew
  errors!eInstance = instance
  errors!eDate = Now
  'errors!eUser = GetUsername
  errors!eNumber = errNumber
  errors!eSource = "HEMS Econ"
  errors!eLine = 0
  errors!eDescription = Left(errDescription, 255)
  errors!eProcedure = errProcedure
  errors!eModule = errModule
  errors.Update
  err.Raise errNumber, "HEMS Econ", errDescription
End Sub

Public Sub Done()
  mInstance = 0
End Sub
