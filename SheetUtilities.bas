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

Public Function getPastedRecordsRng() As Range
    Set getPastedRecordsRng = getRng("Interface", "A9", "L9")
End Function

Public Function getTotalsRng() As Range
    Set getTotalsRng = getRng("Interface", "N2", "Q2")
End Function

Public Function getFinalReportRng() As Range
    Set getFinalReportRng = getRng("Final Report", "A2", "M2")
End Function
Public Function getAddressServiceHeaderRng() As Range
    Dim addressWkSht As Worksheet
    Set addressWkSht = ActiveWorkbook.Worksheets.[_Default]("Addresses")
    Set getAddressServiceHeaderRng = addressWkSht.Range("L1", addressWkSht.Range("L1").End(xlToRight))
End Function
Public Function getAddressesRng() As Range
    Dim lastCol As String
    lastCol = ActiveWorkbook.Worksheets.[_Default]("Addresses").Range("A1").End(xlToRight).Address
    Set getAddressesRng = getRng("Addresses", "A2", lastCol)
End Function

Public Function getDiscardsRng() As Range
    Set getDiscardsRng = getRng("Invalid Discards", "A2", "N2")
End Function

Public Function getAutocorrectRng() As Range
    Set getAutocorrectRng = getRng("Autocorrected Addresses", "A2", "O2")
End Function

Public Function sheetToCSVArray(ByVal sheetName As String, Optional ByVal rng As Range = Nothing) As String()
    ' From https://stackoverflow.com/a/37038840/13342792
    Dim CurrentWB As Workbook
     
    Set CurrentWB = ActiveWorkbook
    
    If rng Is Nothing Then
        ActiveWorkbook.Worksheets.[_Default](sheetName).UsedRange.Copy
    Else
        rng.Copy
    End If
    
    Dim TempWB As Workbook
    Set TempWB = Application.Workbooks.Add(1)
    TempWB.Sheets.[_Default](1).Range("A1").PasteSpecial xlPasteValues
    
    Dim MyFileName As String
    MyFileName = CurrentWB.path & "\test_" & sheetName & Format$(Time, "hh-mm-ss") & ".csv"
    
    Application.DisplayAlerts = False
    TempWB.SaveAs Filename:=MyFileName, FileFormat:=xlCSV, CreateBackup:=False, Local:=True
    TempWB.Close SaveChanges:=False
    Application.DisplayAlerts = True
    
    
    sheetToCSVArray = getCSV(MyFileName)
    Kill (MyFileName)
End Function

Public Sub ClearAll()
    getPastedRecordsRng.Clear
    getTotalsRng.Value = 0
    getFinalReportRng.Clear
    getAddressesRng.Clear
    getAddressServiceHeaderRng.Clear
    getDiscardsRng.Clear
    getAutocorrectRng.Clear
End Sub
