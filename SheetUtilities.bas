Attribute VB_Name = "SheetUtilities"
'@Folder("City_Grant_Address_Report.src")
Option Explicit

Public Function getRng(ByVal sheetName As String, ByVal firstCell As String, ByVal lastCol As String) As Range
    Dim sheet As Worksheet
    Set sheet = ActiveWorkbook.Worksheets.[_Default](sheetName)
        
    Dim lastColNum As Long
    lastColNum = sheet.Range(lastCol).Column
    
    Dim lastRow As Long
    lastRow = sheet.Range(firstCell).End(xlDown).Row
    
    Set getRng = sheet.Range(sheet.Range(firstCell), sheet.Cells.Item(lastRow, lastColNum))
End Function

Public Function getInterfaceRng() As Range
    Set getInterfaceRng = getRng("Interface", "A9", "L9")
End Function

Public Function getFinalReportRng() As Range
    Set getFinalReportRng = getRng("Final Report", "A2", "M2")
End Function

Public Function getAddressesRng() As Range
    Dim lastCol As String
    lastCol = ActiveWorkbook.Worksheets.[_Default]("Addresses").Range("A1").End(xlToRight).Address
    Set getAddressesRng = getRng("Addresses", "A2", lastCol)
End Function

Public Function getDiscardsRng() As Range
    Set getDiscardsRng = getRng("Invalid Discards", "A2", "L2")
End Function

Public Function getAutocorrectRng() As Range
    Set getAutocorrectRng = getRng("Autocorrected Addresses", "A2", "M2")
End Function

Public Function sheetToCSVArray(ByVal sheetName As String) As String()
    Dim CurrentWB As Workbook
     
    Set CurrentWB = ActiveWorkbook
    ActiveWorkbook.Worksheets.[_Default](sheetName).UsedRange.Copy
    
    Dim TempWB As Workbook
    Set TempWB = Application.Workbooks.Add(1)
    TempWB.Sheets.[_Default](1).Range("A1").PasteSpecial xlPasteValues
    
    Dim MyFileName As String
    MyFileName = CurrentWB.path & "\temp.csv"
    
    Application.DisplayAlerts = False
    TempWB.SaveAs Filename:=MyFileName, FileFormat:=xlCSV, CreateBackup:=False, Local:=True
    TempWB.Close SaveChanges:=False
    Application.DisplayAlerts = True
    
    
    sheetToCSVArray = getCSV(MyFileName)
    Kill (MyFileName)
End Function
