VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSalaryStatement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Salary"
Option Compare Database
Option Explicit

Private Const ErrBase As Long = 700
Private mEmpId As Long
Private mSalaryYear As Long
Private mSalaryMonth As Long
Private Salary As New CSalary
Private mEmployee As New CEmployee

Private mExcel As CExcel

Public Sub CreateSalaryStatement(inEmpID As Long, inSalaryYear As Long, inSalaryMonth As Long)
  mEmpId = inEmpID
  mEmployee.EmpID inEmpID
  mEmployee.Init
  Dim mTemplate As String
  mTemplate = Settings.StatementTemplate
  
  mSalaryMonth = inSalaryMonth
  mSalaryYear = inSalaryYear
  
  Set mExcel = New CExcel
  mExcel.openExcel mTemplate
On Error GoTo fileErr
    DoSalary
    AddTextData
    AddSalary
    GoTo finally
fileErr:
  mExcel.CloseWB
  mErr.Raise ErrBase + 1, "Failed to populate excel file", "CreateSalaryStatement", "CSalaryStatement"
finally:
End Sub

Public Sub Save()
  'SaveSalary
  SaveStatement
End Sub

Public Sub CloseMe()
  mExcel.CloseWB
End Sub

Public Sub SaveStatement()
  mExcel.SaveAs BuildFileName()
End Sub

Private Sub AddSalary()
  AddSalaryPayout
  AddSalaryItems
  AddSalaryAccumulates
End Sub

Private Sub AddTextData()
  AddStatementMonth
  AddEmployeeDetails
End Sub

Private Function BuildFileName() As String
  Dim StatementFilename As String, s As String
  s = mEmployee.StatementFolder & mSalaryYear & "\"
  StatementFilename = s & mSalaryYear & "-" & Format$(mSalaryMonth, "00") & "-" & Format$(mEmployee.PayDay, "00") & "-" & mEmployee.PopularName & "-Lönespec"
  BuildFileName = StatementFilename
End Function

Private Sub AddStatementMonth()
  mExcel.WriteCell 2, 9, "Månad " & dbReader.MonthName(mSalaryMonth) & " " & mSalaryYear
End Sub

Private Sub AddEmployeeDetails()
  mExcel.WriteCell 6, 9, mEmployee.FullName
  mExcel.WriteCell 7, 9, mEmployee.Address
  mExcel.WriteCell 8, 9, dbReader.EmpZipCity(mEmpId)
  mExcel.WriteCell 11, 9, mSalaryYear & "-" & Format$(mSalaryMonth, "00") & "-" & Format$(dbReader.PayDay(mEmpId), "00")
  mExcel.Cell(12, 9) = mSalaryYear & " " & dbReader.MonthName(Salary.WorkMonth)
  mExcel.Cell(13, 6) = dbReader.Bank(mEmpId)
  mExcel.Cell(13, 9) = dbReader.BankAccount(mEmpId)
End Sub

Private Sub AddSalaryAccumulates()
  Dim s As New CSalaryAcc
  s.Accumulates mEmpId, mSalaryYear, mSalaryMonth
  With mExcel
    .Cell(72, 6) = s.AccGross
    ' .Cell() = s.AccSocialFee Not yet used
    .Cell(73, 6) = s.AccTax
  End With
  
End Sub

Private Sub AddSalaryItems()
  Dim Items As New CSalaryItems
  Dim row As Long
  row = 18
  If Items.RetrieveItems(mEmpId, Salary.WorkYear, Salary.WorkMonth) > 0 Then
    Do
      With mExcel
        .Cell(row, 2) = Items.ItemDate
        .Cell(row, 3) = "K/" & Items.CustomerPopularName
        .Cell(row, 6) = Items.WorkedHours
        .Cell(row, 10) = Items.Price
        .Cell(row, 11) = Items.Pay
        .Cell(row, 9) = "timmar"
        row = row + 1
      End With
    Loop Until Items.NextItem
  End If
End Sub

Private Sub AddSalaryPayout()
  Const offset = 30 'After adjusting excel template 2022-11-16 to fit more salary items
  mExcel.Cell(42 + offset, 11) = Salary.GrossSalary
  mExcel.Cell(43 + offset, 11) = Salary.Tax
  mExcel.Cell(46 + offset, 11) = Salary.NetSalary
  mExcel.Cell(45 + offset, 3) = Salary.SocialFee
  mExcel.Cell(44 + offset, 3) = Salary.SocialFeeRate * 100 'To percent
  mExcel.Cell(42 + offset, 3) = Salary.TaxTable
  mExcel.Cell(29 + offset, 11) = Salary.VacationAllowance
  mExcel.Cell(30 + offset, 6) = Salary.WorkedHours
End Sub

Private Sub DoSalary()
  Salary.EmpID = mEmpId
  Salary.SalaryMonth = mSalaryMonth
  Salary.SalaryYear = mSalaryYear
  Salary.Init
End Sub

Private Function SalaryStored() As Boolean
  SalaryStored = dbReader.SalaryRegistered(mEmpId, mSalaryYear, mSalaryMonth)
End Function

Private Sub SaveSalary()
  dbReader.SaveSalary mEmpId, mSalaryYear, mSalaryYear
End Sub

