Attribute VB_Name = "dbReader"
'@Folder "HEMS"
Option Compare Database
Option Explicit

Public Property Get MonthName(inMonth As Long) As String
  Dim rst As DAO.Recordset
  Set rst = CurrentDb.OpenRecordset("SELECT MonthName FROM Months WHERE MonthNr = " & inMonth)
  MonthName = rst!MonthName
End Property

Public Property Get EmpFullname(inEmpID As Long) As String
  Dim rst As DAO.Recordset
  Set rst = CurrentDb.OpenRecordset("SELECT FirstName, LastName FROM Employees WHERE ID = " & inEmpID)
  EmpFullname = rst!FirstName & " " & rst!LastName
  CurrentDb.Close
End Property

Public Property Get EmpPopularname(inEmpID As Long) As String
  Dim rst As DAO.Recordset
  Set rst = CurrentDb.OpenRecordset("SELECT PopularName FROM Employees WHERE ID = " & inEmpID)
  EmpPopularname = rst!PopularName
  CurrentDb.Close
End Property

Public Property Get EmpAddress(inEmpID As Long) As String
  Dim rst As DAO.Recordset
  Set rst = CurrentDb.OpenRecordset("SELECT Street, StreetNumber, AptNr FROM Employees WHERE ID = " & inEmpID)
  EmpAddress = rst!Street & " " & rst!StreetNumber & " " & rst!AptNr
  CurrentDb.Close
End Property

Public Property Get EmpZipCity(inEmpID As Long) As String
  Dim rst As DAO.Recordset
  Set rst = CurrentDb.OpenRecordset("SELECT Zip, City FROM Employees WHERE ID = " & inEmpID)
  EmpZipCity = rst!Zip & " " & rst!City
  CurrentDb.Close
End Property

Public Property Get HourlyPay(inEmpID As Long) As Currency
  Dim rst As DAO.Recordset
  Set rst = CurrentDb.OpenRecordset("SELECT SalaryPerHour FROM Employees WHERE ID = " & inEmpID)
  HourlyPay = rst!SalaryPerHour
  CurrentDb.Close
End Property

Public Property Get PayDay(inEmpID As Long) As Long
  Dim rst As DAO.Recordset
  Set rst = CurrentDb.OpenRecordset("SELECT PayDay FROM Employees WHERE ID = " & inEmpID)
  PayDay = rst!PayDay
  CurrentDb.Close
End Property

Public Property Get Salary(inEmpID As Long, inYear As Long, inMonth As Long) As CSalary
  Set Salary = New CSalary
  Salary.EmpID = inEmpID
  Salary.SalaryMonth = inMonth
  Salary.SalaryYear = inYear
  Salary.Init
End Property

Public Property Get Bank(inEmpID As Long) As String
  Dim rst As DAO.Recordset
  Set rst = CurrentDb.OpenRecordset("SELECT Bank FROM Employees WHERE ID = " & inEmpID)
  Bank = rst!Bank
  CurrentDb.Close
End Property

Public Property Get BankAccount(inEmpID As Long) As String
  Dim rst As DAO.Recordset
  Set rst = CurrentDb.OpenRecordset("SELECT Account FROM Employees WHERE ID = " & inEmpID)
  BankAccount = rst!Account
  CurrentDb.Close
End Property

Public Property Get CustomerPopularName(inCustomerID As Long) As String
  Dim rst As DAO.Recordset
  Set rst = CurrentDb.OpenRecordset("SELECT PopularName FROM Customers WHERE ID = " & inCustomerID)
  CustomerPopularName = rst!PopularName
  CurrentDb.Close
End Property

Public Property Get StatementFolder(inEmpID As Long) As String
  Dim rst As DAO.Recordset
  Set rst = CurrentDb.OpenRecordset("SELECT StatementFolder FROM Employees WHERE ID = " & inEmpID)
  StatementFolder = rst!StatementFolder
  CurrentDb.Close
End Property

Public Function RetrieveSalary(inEmpID As Long, inYear As Long, inMonth As Long) As DAO.Recordset
  Dim rst As DAO.Recordset
  Dim sql  As String
  sql = "SELECT * FROM SalaryPayOuts WHERE EmployeeID = " & inEmpID & " PayYear = " & inYear & " PayMonth = " & inMonth
  Set rst = CurrentDb.OpenRecordset(sql)
  Set RetrieveSalary = rst
End Function

Public Function SalaryRegistered(inEmpID As Long, inYear As Long, inMonth As Long) As Boolean
  Dim rst As DAO.Recordset
  Set rst = RetrieveSalary(inEmpID, inYear, inMonth)
  SalaryRegistered = False
  If rst.RecordCount > 0 Then SalaryRegistered = True
End Function

Public Sub SaveSalary(inEmpID As Long, inYear As Long, inMonth As Long)
  Dim rst As DAO.Recordset
  Set rst = CurrentDb.OpenRecordset("SalaryPayOuts", dbOpenDynaset, dbSeeChanges)

  rst.AddNew
  rst!EmployeeID = inEmpID
  rst!PayDay = PayDay(inEmpID)
  rst!PayMonth = inMonth
  rst!PayYear = inYear
  rst!gross = Salary(inEmpID, inYear, inMonth).GrossSalary
  rst!Tax = Salary(inEmpID, inYear, inMonth).Tax
  rst!SocialFee = Salary(inEmpID, inYear, inMonth).SocialFee
  
  rst.Update
  CurrentDb.Close
End Sub

