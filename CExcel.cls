VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Utilities")
Option Compare Database
Option Explicit
Const ExcelOffTheGrid As String = "https://exceloffthegrid.com/vba-code-save-excel-file-as-pdf/"

Private wb As Workbook

Public Sub openExcel(ByVal inFileName As String)
  Dim appExcel As excel.Application
  
  Set appExcel = CreateObject("Excel.Application")
  Set wb = appExcel.Workbooks.Open(inFileName, , True)
  appExcel.Visible = True
  
End Sub

Public Sub Save()
  wb.Save
End Sub

Public Sub SaveAs(inFileName As String)
  Dim alerts As Boolean
  mDirectoryFunctions.CreateDirTree inFileName
  alerts = wb.Application.DisplayAlerts
  wb.Application.DisplayAlerts = False
    wb.SaveAs FileName:=inFileName, FileFormat:=1, ConflictResolution:=xlLocalSessionChanges
    wb.ExportAsFixedFormat Type:=xlTypePDF, FileName:=inFileName
  wb.Application.DisplayAlerts = alerts
  Application.FollowHyperlink mDirectoryFunctions.StripFilename(inFileName)
End Sub

Public Sub CloseWB()
  wb.Close False                                 'Don't save changes
  Set wb = Nothing
End Sub

Public Sub WriteCell(ByVal inRow As Long, ByVal inCol As Long, ByVal value As Variant)
  wb.Application.Cells.Item(inRow, inCol) = value
End Sub

Public Property Let Cell(inRow As Long, inCol As Long, inVal As Variant)
  WriteCell inRow, inCol, inVal
End Property

Public Property Get Cell(inRow As Long, inCol As Long) As Variant
  Cell = wb.Application.Cells(inRow, inCol)
End Property

Public Property Let Border(inRow As Long, inCol As Long, inBorder As Long, inStyle As Long)
  Dim adr As String
  adr = wb.Application.Cells.Item(inRow, inCol).Address
  wb.Application.Range(adr).Borders(inBorder).LineStyle = inStyle
  'wb.Application.Range(Chr(inCol + 64) & inRow).Borders(inBorder).LineStyle = inStyle
End Property



'Save active sheet as PDF
'The following code saves the selected sheets as a single PDF.

Public Sub SaveActiveSheetsAsPDF()

  'Create and assign variables
  Dim saveLocation As String
  saveLocation = "C:\Users\marks\OneDrive\Documents\myPDFFile.pdf"

  'Save Active Sheet(s) as PDF
  ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
                                  FileName:=saveLocation

End Sub

'Save active workbook as PDF
'To save all the sheets from a workbook use the following macro.

Public Sub SaveActiveWorkbookAsPDF()

  'Create and assign variables
  Dim saveLocation As String
  saveLocation = "C:\Users\marks\OneDrive\Documents\myPDFFile.pdf"

  'Save active workbook as PDF
  ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, _
                                     FileName:=saveLocation

End Sub


