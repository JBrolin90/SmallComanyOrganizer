VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSugeySalary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("SugeySalary")
Option Compare Database
Option Explicit



Private Sales As New CMonthlySales
Private salaries As New CMonthlySalary
Private expenses As New CExpenses


Public Sub Init(Year As Long, Month As Long)
  Sales.SumTotals Month, Year
  salaries.SumTotals Month, Year
  
End Sub

Public Property Get GrossSalary()
  Dim profit As Double
  profit = Sales.CustomerSales + Sales.RUT - salaries.Total - expenses.netTotal
  If profit <= 0 Then
    GrossSalary = 0
  Else
    GrossSalary = profit - mTax.Egenavgift(profit)
  End If
End Property



Public Property Get Tax()
  If GrossSalary <= 0 Then
    Tax = 0
  Else
    Tax = mTax.ASkatt(GrossSalary, 31)  'Sugey has table 31
  End If
End Property

Public Property Get NetSalary()
  If GrossSalary <= 0 Then
    NetSalary = 0
  Else
    NetSalary = GrossSalary - Tax
  End If
End Property
