VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Invoice"
Option Compare Database
Option Explicit

Private Const ErrBase As Long = 800

Private mMonth As Long
Private mYear As Long

Private mCustomer As New CCustomer
Private mExcel As New CExcel
Private mCalc As New CInvoiceCalc


Public Sub CreateInvoice(inCustomerID As Long, inYear As Long, inMonth As Long)
  mCustomer.CustomerID inCustomerID
  mCalc.Calculate inCustomerID, inYear, inMonth
  mMonth = inMonth
  mYear = inYear
  
  mExcel.openExcel Settings.InvoiceTemplate
  On Error GoTo excelError
    EmitTheNumbers
    FillCustomerData
    EmitPeriod
    GoTo finally
excelError:
  mExcel.CloseWB
  mErr.Raise ErrBase + 1, "Failed to populate Invoice", "CreateInvoice", "CInvoice"
finally:
    
End Sub

Public Sub Save()
  mExcel.SaveAs BuildFileName()
End Sub

Public Sub CloseMe()
  mExcel.CloseWB
End Sub

Private Function BuildFileName() As String
  Dim s As String
  Dim d As String
  d = DateSerial(mYear, mMonth, 1)
  s = mCustomer.InvoiceFolder & mYear & "\"
  BuildFileName = s & mYear & "-" & Format$(mMonth, "00") & "-" & Format$(Day(mDateFunctions.GetLastDayOfMonth(d)), "00") & "-" & mCustomer.PopularName & "-Faktura"
End Function


Private Sub EmitPeriod()
  Dim s As String
  s = mExcel.Cell(13, 2)
  s = Replace(s, "<Month>", dbReader.MonthName(mMonth))
  s = Replace(s, "<Year>", mYear)
  mExcel.Cell(13, 2) = s
End Sub


Private Sub EmitTheNumbers()
  With mExcel
    .Cell(24, 10) = mCalc.VAT
    If mCustomer.RUT Then
      .Cell(25, 10) = mCalc.AmountTotal * 2
      .Cell(27, 10) = mCalc.AmountTotal
      .Cell(28, 10) = mCalc.AmountTotal
      .Cell(50, 4) = mCalc.AmountTotal
      .Cell(24, 7) = "Moms ingår med"
      .Cell(24, 5) = ""
    Else
      .Cell(25, 10) = mCalc.AmountTotal
      .Cell(27, 10) = ""
      .Cell(28, 10) = mCalc.AmountTotal + mCalc.VAT
      .Cell(50, 4) = mCalc.AmountTotal + mCalc.VAT
      .Cell(24, 7) = "Moms"
      .Cell(24, 5) = ""
      
    End If
  End With
  EmitItems
End Sub

Private Sub EmitItems()
  Dim i As Long
  Dim topRow As Long
  i = 0
  topRow = 15
  mCalc.Beginning
  Do
    mExcel.Cell(i + topRow, 2) = mCalc.ItemJob
    mExcel.Cell(i + topRow, 4) = mCalc.ItemDate
    mExcel.Cell(i + topRow, 5) = mCalc.ItemStart
    mExcel.Cell(i + topRow, 6) = mCalc.ItemEnd
    mExcel.Cell(i + topRow, 7) = mCalc.CurrentHours
    If mCustomer.RUT Then
      mExcel.Cell(i + topRow, 8) = mCalc.CurrentPrice * 2
      mExcel.Cell(i + topRow, 10) = mCalc.CurrentHours * mCalc.CurrentPrice * 2
    Else
      mExcel.Cell(i + topRow, 8) = mCalc.CurrentPrice
      mExcel.Cell(i + topRow, 10) = mCalc.CurrentHours * mCalc.CurrentPrice
    End If
    i = i + 1
  Loop While mCalc.NextItem And i < 9
  If i > 9 Then err.Raise 1, "CInvoice.EmitItems", "För många rader, Max 9 är tillåtet"
End Sub


Private Sub FillCustomerData()
  Dim d As Date
  d = DateSerial(mYear, mMonth, 1)

  With mExcel
    .Cell(4, 8) = mYear & mMonth & mCustomer.getCustomerID() & "-1" '"Invoice number"
    .Cell(38, 3) = .Cell(4, 8)
    .Cell(4, 10) = mYear & "-" & Format$(mMonth, "00") & "-" & Format$(Day(mDateFunctions.GetLastDayOfMonth(d)), "00") '"Invoice Date"
    .Cell(37, 3) = .Cell(4, 10)
    .Cell(8, 3) = DateAdd("d", 30, .Cell(4, 10))         '"Due date"
    .Cell(39, 3) = .Cell(8, 3)
    
    With mCustomer
      mExcel.Cell(7, 8) = .FullName '"Full name"
      mExcel.Cell(45, 2) = .FullName '"Full name"
      mExcel.Cell(8, 8) = .Address '"Address"
      mExcel.Cell(46, 2) = .Address '"Address"
      mExcel.Cell(9, 8) = .Zip '"Zip"
      mExcel.Cell(47, 2) = .Zip '"Zip"
      mExcel.Cell(9, 9) = .City '"City"
      mExcel.Cell(47, 3) = .City '"City"
    End With
    
    .Cell(10, 3) = mCustomer.FullName '"Your reference"
    .Cell(11, 3) = "Sugey Brolin"
    .Cell(10, 9) = mCustomer.getCustomerID() '"Customer number"
  End With
  FillSKVData
End Sub

Private Sub FillSKVData()
  Dim s As String
  mExcel.Cell(27, 5) = ""
  mExcel.Cell(28, 7) = "Att betala"
  If mCustomer.RUT Then
    s = mExcel.Cell(28, 2)
    s = Replace(s, "<ADDRESS>", mCustomer.FullAddress)
    s = Replace(s, "<PERSONNUMMER>", mCustomer.PersonalNumber)
    mExcel.Cell(27, 7) = "Avgår skattereduktion"
  Else
   s = ""
  End If
  mExcel.Cell(28, 2) = s
End Sub

