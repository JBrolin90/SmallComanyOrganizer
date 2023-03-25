VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMonthlySalary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Monthly")
Option Compare Database
Option Explicit

Private mGross As Currency
Private mATax As Currency
Private mSocial As Currency
Private mNet As Currency
Private mTotal As Currency


Public Sub SumTotals(inMonth As Long, inYear As Long)
  mGross = 0
  mATax = 0
  mSocial = 0
  mNet = 0
  mTotal = 0
  Dim employees As New CEmployee
  employees.EmpID
  employees.Begin
  Dim monthlySalaries As New CSalary
  monthlySalaries.SalaryMonth = inMonth
  monthlySalaries.SalaryYear = inYear
  Debug.Print
  Debug.Print "== BEGIN SALARIES =="
  Do
    monthlySalaries.EmpID = employees.getEmpID()
    monthlySalaries.Init
    mGross = mGross + monthlySalaries.GrossSalary
    mATax = mATax + monthlySalaries.Tax
    mSocial = mSocial + monthlySalaries.SocialFee
    mNet = mNet + monthlySalaries.NetSalary
    mTotal = mGross + mSocial
    If monthlySalaries.GrossSalary > 0 Then
      Debug.Print employees.FullName & " " & employees.getEmpID() & " " & monthlySalaries.GrossSalary
    End If
    DoEvents
  Loop While employees.nextEmployee
  Debug.Print "mGross=" & mGross & " mATax=" & mATax & " mSocial" & mSocial & " mNet=" & mNet & " mTotal=" & mTotal
  Debug.Print "== END SALARIES =="
End Sub

Public Property Get GrossSalary() As Currency
  GrossSalary = mGross
End Property

Public Property Get Tax() As Currency
  Tax = mATax
End Property

Public Property Get Social() As Currency
  Social = mSocial
End Property

Public Property Get Net() As Currency
  Net = mNet
End Property

Public Property Get Total() As Currency
  Total = mTotal
End Property

Public Property Get TotalTax() As Currency
  TotalTax = Tax + Social
End Property

