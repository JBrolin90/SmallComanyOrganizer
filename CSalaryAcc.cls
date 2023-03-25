VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSalaryAcc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Salary"
Option Compare Database
Option Explicit

Private mGross As Long
Private mNet As Long
Private mTax As Long
Private mSoc As Long

Public Sub Accumulates(inEmpID As Long, inYear As Long, inUntilMonth As Long)
  Dim s As New CSalary, i As Integer
  mGross = 0
  mNet = 0
  mTax = 0
  mSoc = 0
  
  For i = 1 To inUntilMonth
    s.EmpID = inEmpID
    s.SalaryYear = inYear
    s.SalaryMonth = i
    s.Init
    mGross = mGross + s.GrossSalary
    mTax = mTax + s.Tax
    mNet = mNet + s.NetSalary
    mSoc = mSoc + s.SocialFee
  Next i
  
End Sub

Public Property Get AccSocialFee() As Long
  AccSocialFee = mSoc
End Property

Public Property Get AccTax() As Long
  AccTax = mTax
End Property

Public Property Get AccGross() As Long
  AccGross = mGross
End Property

Public Property Get AccNet() As Long
  AccNet = mNet
End Property

