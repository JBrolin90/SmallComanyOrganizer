VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMonthlySales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Monthly")
Option Compare Database
Option Explicit

'@Ignore MoveFieldCloserToUsage
Private mSales As Currency
Private mVAT As Currency
Private mRUT As Currency
Private mNet As Currency
Private mTotal As Currency


Public Sub SumTotals(inMonth As Long, inYear As Long)
  mSales = 0
  mVAT = 0
  mNet = 0
  mRUT = 0
  mTotal = 0
  Dim customers As New CCustomer
  customers.CustomerID
  customers.Begin
  Debug.Print
  Debug.Print "=== START INVOICES ===="
  Do
    Dim monthlySales As New CInvoiceCalc
    monthlySales.Calculate customers.getCustomerID(), inYear, inMonth
    mSales = mSales + monthlySales.AmountTotal
    mVAT = mVAT + monthlySales.VAT
    mRUT = mRUT + monthlySales.RUT
    mNet = mNet + monthlySales.Net
    mTotal = mNet + mVAT + mRUT
    If monthlySales.Net > 0 Then
      Debug.Print customers.FullName & " " & customers.getCustomerID() & " " & monthlySales.AmountTotal
    End If
    DoEvents
  Loop While customers.nextCustomer
  Debug.Print "TOTALS: mVAT=" & mVAT & " mRUT=" & mRUT & " mNet=" & mNet & "mTotal=" & mTotal
  Debug.Print "=== END INVOICES ===="
End Sub

Public Property Get CustomerSales() As Currency
  CustomerSales = mNet
End Property

Public Property Get VAT() As Currency
  VAT = mVAT
End Property

Public Property Get RUT() As Currency
  RUT = mRUT
End Property

Public Property Get Sales()
  Sales = RUT + CustomerSales
End Property

Public Property Get Total() As Currency
  Total = mTotal
End Property

