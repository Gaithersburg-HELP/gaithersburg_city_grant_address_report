Attribute VB_Name = "SheetUtilities"
'@Folder("City_Grant_Address_Report.src")
Option Explicit
' Returns blank row after all data
Public Function getBlankRow(ByVal sheetName As String) As Range
    Dim sheet As Worksheet
    Set sheet = ActiveWorkbook.Worksheets.[_Default](sheetName)
    
    Set getBlankRow = sheet.Rows.Item(sheet.Rows.Item(sheet.Rows.Count).End(xlUp).row + 1)
End Function

' Returns all data below (all cells between firstCell and lastCol) including blanks and firstCell
Public Function getRng(ByVal sheetName As String, ByVal firstCell As String, ByVal lastCol As String) As Range
    Dim sheet As Worksheet
    Set sheet = ActiveWorkbook.Worksheets.[_Default](sheetName)
        
    Dim lastColNum As Long
    lastColNum = sheet.Range(lastCol).Column
    
    Dim lastRow As Long
    lastRow = sheet.Range(firstCell).row
    
    Dim i As Long
    i = sheet.Range(firstCell).Column
    Do While i <= lastColNum
        Dim currentLastRow As Long
        currentLastRow = sheet.Cells.Item(sheet.Rows.Count, i).End(xlUp).row
        If (currentLastRow > lastRow) Then lastRow = currentLastRow
        i = i + 1
    Loop
    
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

Private Function getServiceHeaderLastCell(ByVal sheetName As String, ByVal cellToLeftOfHeaders As String) As String
    getServiceHeaderLastCell = ActiveWorkbook.Worksheets.[_Default](sheetName) _
                                      .Range(cellToLeftOfHeaders).End(xlToRight).address
End Function

Public Function getServiceHeaderRng(ByVal sheetName As String) As Range
    Set getServiceHeaderRng = ActiveWorkbook.Worksheets.[_Default](sheetName) _
                                    .Range("O1:" & getServiceHeaderLastCell(sheetName, "N1"))
End Function

Public Function getAddressRng(ByVal sheetName As String) As Range
    Set getAddressRng = getRng(sheetName, "A2", "M2")
End Function

Public Function getAddressVisitDataRng(ByVal sheetName As String) As Range
    Set getAddressVisitDataRng = Application.Union(getRng(sheetName, "N2", "N2"), _
                                                   getRng(sheetName, "N2", _
                                                          getServiceHeaderLastCell(sheetName, "N1")), _
                                                   getServiceHeaderRng(sheetName))
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
    
    Dim i As Long
    For i = 3 To ActiveWorkbook.Sheets.Count
        getAddressRng(ActiveWorkbook.Sheets.[_Default](i).Name).Clear
        getAddressVisitDataRng(ActiveWorkbook.Sheets.[_Default](i).Name).Clear
        getServiceHeaderRng(ActiveWorkbook.Sheets.[_Default](i).Name).Clear
    Next
End Sub


