VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CVAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Monthly")
Option Compare Database
Option Explicit

Private Const sqlVAT As String = " SELECT SUM(DurationPlanned * PricePerHour) as itemPrice FROM Schedules "
Private Const sqlPeriod As String = " WHERE Month(StartTimePlanned) Like <parMonth> AND Year(StartTimePlanned) Like <parYear> "

Public Function VATMonth(inYear As Long, inMonth As Long) As Currency
  Dim sql As String
  sql = sqlVAT & Replace(Replace(sqlPeriod, "<parMonth>", inMonth), "<parYear>", inYear)
  VATMonth = db.GetRecordset(sql)!itemPrice
End Function
