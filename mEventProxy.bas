Attribute VB_Name = "mEventProxy"
'@Folder("HEMS")
Option Compare Database
Option Explicit

Private proxy As CEventProxy

Public Property Get EventProxy() As CEventProxy
  If proxy Is Nothing Then
    Set proxy = New CEventProxy
  End If
  Set EventProxy = proxy
End Property

