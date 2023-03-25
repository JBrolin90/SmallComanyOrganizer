Attribute VB_Name = "Settings"
'@Folder "HEMS"
Option Compare Database
Option Explicit

Private Const ErrBase As Long = 500

Public Property Get Version() As String
  Version = Nz(Tag("Version"), "1.0")
End Property

Public Property Get Templates() As String
  If Version = "1.0" Then
    Templates = BaseFolder
  ElseIf Version = "2.0" Then
    Templates = HemsEconFolder + Tag("Templates")
  Else
    Raise ErrBase + 1, "Version " + Version + " Not supported", "Templates", "Settings"
  End If
End Property

Public Property Get StatementTemplate() As String
  If Version = "1.0" Then
    StatementTemplate = Tag("StatementTemplate")
  ElseIf Version = "2.0" Then
    StatementTemplate = Templates + Tag("StatementTemplate")
  Else
    Raise ErrBase + 2, "Version " + Version + " Not supported", "StatementTemplate", "Settings"
  End If
End Property

Public Property Get SchemaTemplate() As String
  If Version = "1.0" Then
    SchemaTemplate = Tag("SchemaTemplate")
  ElseIf Version = "2.0" Then
    SchemaTemplate = Templates + Tag("SchemaTemplate")
  Else
    Raise ErrBase + 3, "Version " + Version + " Not supported", "SchemaTemplate", "Settings"
  End If
End Property

Public Property Get InvoiceTemplate() As String
  InvoiceTemplate = Tag("InvoiceTemplate")
  If Version = "1.0" Then
    InvoiceTemplate = Tag("InvoiceTemplate")
  ElseIf Version = "2.0" Then
    InvoiceTemplate = Templates + Tag("InvoiceTemplate")
  Else
    Raise ErrBase + 4, "Version " + Version + " Not supported", "InvoiceTemplate", "Settings"
  End If
End Property

Public Property Get HemsEconFolder() As String
  If Version = "2.0" Then
    HemsEconFolder = Tag("HEMS Econ Folder")
  Else
    Raise ErrBase + 1, "Version " + Version + " Not supported", "HemsEconFolder", "Settings"
  End If
End Property

Public Property Get BaseFolder() As String
  If Version = "1.0" Then
    BaseFolder = Tag("BaseFolder")
  ElseIf Version = "2.0" Then
    BaseFolder = Tag("BaseFolder")
  Else
    Raise ErrBase + 1, "Version " + Version + " Not supported", "BaseFolder", "Settings"
  End If
End Property
Public Property Get SchemaBaseFolder() As String
  If Version = "1.0" Then
    SchemaBaseFolder = Tag("SchemaBaseFolder")
  ElseIf Version = "2.0" Then
    SchemaBaseFolder = BaseFolder + Tag("SchemaBaseFolder")
  Else
    Raise ErrBase + 1, "Version " + Version + " Not supported", "SchemaBaseFolder", "Settings"
  End If
End Property

Public Property Get InvoiceBaseFolder() As String
  If Version = "1.0" Then
    InvoiceBaseFolder = Tag("InvoiceBaseFolder")
  ElseIf Version = "2.0" Then
    InvoiceBaseFolder = BaseFolder + Tag("InvoiceBaseFolder")
  Else
    Raise ErrBase + 1, "Version " + Version + " Not supported", "InvoiceBaseFolder", "Settings"
  End If
End Property

Public Property Get StatmentBaseFolder() As String
  If Version = "1.0" Then
    StatmentBaseFolder = Tag("StatementBaseFolder")
  ElseIf Version = "2.0" Then
    StatmentBaseFolder = BaseFolder + Tag("StatementBaseFolder")
  Else
    Raise ErrBase + 1, "Version " + Version + " Not supported", "InvoiceBaseFolder", "Settings"
  End If
End Property

Public Property Get Tag(inTag As String) As String
  Dim rst As DAO.Recordset
  Dim s As String
  s = "SELECT TagValue FROM LocalSettings WHERE Tag = " & "'" & inTag & "'"
  Set rst = CurrentDb.OpenRecordset(s)
  Tag = rst!TagValue
  CurrentDb.Close
End Property

Public Property Let Tag(inTag As String, inTagValue As String)
  Dim rst As DAO.Recordset
  Set rst = CurrentDb.OpenRecordset("SELECT TagValue FROM LocalSettings WHERE Tag = " & inTag)
  rst.Edit
  rst!TagValue = inTagValue
  rst.Update
  CurrentDb.Close
End Property

