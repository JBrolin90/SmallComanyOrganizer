Attribute VB_Name = "mTax"
'@Folder("Salary")
Option Compare Database
Option Explicit

Const SkatterOchAvgifterHandelsbolag As String = "https://www.verksamt.se/starta/skatter-och-avgifter/handelsbolag-eller-kommanditbolag#:~:text=I%20ett%20handelsbolag%20beskattas%20del%C3%A4garna,ut%20pengar%20ur%20ditt%20bolag."


Public Function ASkatt(Salary As Double, TaxTable As Long) As Double
  Dim rst As DAO.Recordset
  Dim sql As String
  Dim qdf As DAO.QueryDef
  sql = "Select * from TaxTable where TabellNr = <table> and " & _
          " [Antal dgr] = '30B' and " & _
          " [inkomst from] <= <salary>  and [inkomst tom] >= <salary>"
  sql = Replace(sql, "<table>", TaxTable, 1, 10)
  sql = Replace(sql, "<salary>", Salary, 1, 10)
  Set rst = db.GetRecordset(sql)
    
  ASkatt = rst![Kolumn 1]
End Function

Public Function Egenavgift(inProfit As Double) As Double
  Dim eAvg As Double
  eAvg = inProfit * 0.75 'Schablonavdrag 25%
  eAvg = eAvg * 0.2897 'Egenavgift 28.97%
  Egenavgift = eAvg
End Function
