VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSalary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Salary"
Option Compare Database
Option Explicit

Private mSalaryMonth As Long
Private mSalaryYear As Long
Private mEmpId As Long
Private mEmployee As CEmployee

Private dbs As DAO.Database
Private rstSalary As DAO.Recordset
Private rstTax As DAO.Recordset

Public Property Let SalaryMonth(inMonth As Long)
  mSalaryMonth = inMonth
End Property

Public Property Get SalaryMonth() As Long
  SalaryMonth = mSalaryMonth
End Property

Public Property Let SalaryYear(inYear As Long)
  mSalaryYear = inYear
End Property

Public Property Get SalaryYear() As Long
  SalaryYear = mSalaryYear
End Property

Public Property Get WorkYear() As Long
  Dim Y As Long
  Y = SalaryYear
  If mEmployee.PayDay > 10 Then
    If WorkMonth = 12 Then Y = SalaryYear - 1
  Else
    If WorkMonth > 10 Then Y = SalaryYear - 1
  End If
  WorkYear = Y
End Property

Public Property Get WorkMonth() As Long
  Dim m As Long
  m = SalaryMonth
  If mEmployee.PayDay > 10 Then
    m = m - 1
  Else
    m = m - 2
  End If
  If m < 1 Then m = m + 12
  WorkMonth = m
End Property


Public Property Let EmpID(inEmpID As Long)
  mEmpId = inEmpID
End Property

Public Property Get EmpID() As Long
  EmpID = mEmpId
End Property

Public Property Get TaxTable() As Long
  TaxTable = mEmployee.TaxTable
End Property

Public Property Get SocialFeeRate() As Double
  SocialFeeRate = mEmployee.SocialFeeRate
End Property

Public Property Get SocialFee() As Long
  SocialFee = GrossSalary * SocialFeeRate
End Property

Public Property Get Employee() As CEmployee
  Set Employee = mEmployee
End Property

Public Sub Init()
  Set dbs = CurrentDb
  Set mEmployee = New CEmployee
  mEmployee.EmpID EmpID
  mEmployee.Init
  RetrieveSalary
  RetrieveTax
End Sub

Public Property Get GrossSalary() As Currency
  GrossSalary = 0
  If rstSalary.RecordCount > 0 Then
    GrossSalary = Nz(rstSalary!Salary, 0)
  End If
End Property

Public Property Get Tax() As Currency
  Tax = 0
  If rstTax.RecordCount > 0 Then
    Tax = rstTax![Kolumn 1]
  End If
End Property

Public Property Get NetSalary() As Currency
  NetSalary = GrossSalary - Tax
End Property

Public Property Get WorkedHours() As Double
  WorkedHours = 0
  If rstSalary.RecordCount > 0 Then
    WorkedHours = rstSalary!hours
  End If
End Property

Public Property Get VacationAllowance() As Currency
  VacationAllowance = GrossSalary / 1.12 * 0.12
End Property

Private Sub RetrieveSalary()
  Dim qdf As DAO.QueryDef

  Set qdf = dbs.QueryDefs("Salary")
  qdf.Parameters("Month") = WorkMonth
  qdf.Parameters("Year") = WorkYear
  qdf.Parameters("EmployeeID") = EmpID
  Set rstSalary = qdf.OpenRecordset()
End Sub

Private Sub RetrieveTax()
  Dim qdf As DAO.QueryDef
  
  Set qdf = dbs.QueryDefs("Tax")
  qdf.Parameters("Month") = WorkMonth
  qdf.Parameters("Year") = WorkYear
  qdf.Parameters("EmployeeID") = EmpID
  qdf.Parameters("Table") = TaxTable
  
  Set rstTax = qdf.OpenRecordset()
End Sub

Private Sub Class_Terminate()
  CurrentDb.Close
End Sub

