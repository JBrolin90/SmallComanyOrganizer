VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLongIterator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Utilities"
Option Compare Database
Option Explicit

Private lStart As Long
Private lEnd As Long
Private lStep As Long
Private cursor As Long

Public Sub Init(inStart As Long, inEnd As Long)
  lStart = inStart
  lEnd = inEnd
  lStep = 1
  Reinit
End Sub

Public Sub Reinit()
  cursor = lStart - lStep
End Sub

Public Property Get NextVal() As Long
  cursor = cursor + lStep
  NextVal = cursor
  Debug.Assert NextVal <= lEnd
End Property

Public Property Get Finished() As Boolean
  Finished = False
  If cursor + lStep > lEnd Then Finished = True
End Property

