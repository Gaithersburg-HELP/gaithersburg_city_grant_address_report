Attribute VB_Name = "SheetUtilities"
'@Folder("City_Grant_Address_Report.src")
Option Explicit

Public Const firstServiceColumn As Long = 16

Public Function serviceFirstCell() As String
    serviceFirstCell = ActiveSheet.Range("A1").offset(0, firstServiceColumn - 1).address
End Function

Public Function rxFirstCell() As String
    rxFirstCell = ActiveSheet.Range("A2").offset(0, firstServiceColumn - 2).address
End Function

' Returns blank row after all data, assuming Column A is filled in last row
Public Function getBlankRow(ByVal sheetName As String) As Range
    Dim sheet As Worksheet
    Set sheet = ActiveWorkbook.Worksheets.[_Default](sheetName)
    
    Set getBlankRow = sheet.rows.Item(sheet.rows.Item(sheet.rows.Count).End(xlUp).row + 1)
End Function

' Returns all data below (all cells between firstCell and lastCol) including blanks and firstCell
Public Function getRng(ByVal sheetName As String, ByVal firstCell As String, ByVal lastCol As String) As Range
    Dim sheet As Worksheet
    Set sheet = ActiveWorkbook.Worksheets.[_Default](sheetName)
        
    Dim lastColNum As Long
    lastColNum = sheet.Range(lastCol).column
    
    Dim lastRow As Long
    lastRow = sheet.Range(firstCell).row
    
    Dim i As Long
    i = sheet.Range(firstCell).column
    Do While i <= lastColNum
        Dim currentLastRow As Long
        currentLastRow = sheet.Cells.Item(sheet.rows.Count, i).End(xlUp).row
        If (currentLastRow > lastRow) Then lastRow = currentLastRow
        i = i + 1
    Loop
    
    Set getRng = sheet.Range(sheet.Range(firstCell), sheet.Cells.Item(lastRow, lastColNum))
End Function

Public Function getPastedRecordsRng() As Range
    Set getPastedRecordsRng = getRng("Interface", "A23", "L23")
End Function

Public Function getTotalsRng() As Range
    Set getTotalsRng = ActiveWorkbook.Worksheets.[_Default]("Interface").Range("N2:Q6")
End Function

Public Function getFinalReportRng() As Range
    Set getFinalReportRng = getRng("Final Report", "A2", "M2")
End Function

Private Function getServiceHeaderLastCell(ByVal sheetName As String) As String
    getServiceHeaderLastCell = ActiveWorkbook.Worksheets.[_Default](sheetName) _
                                      .Range("A1").offset(0, firstServiceColumn - 2) _
                                      .End(xlToRight).address
End Function

Public Function getServiceHeaderRng(ByVal sheetName As String) As Range
    Set getServiceHeaderRng = ActiveWorkbook.Worksheets.[_Default](sheetName) _
                                    .Range(serviceFirstCell, getServiceHeaderLastCell(sheetName))
End Function

' Returns zero based service array
Public Function loadServiceNames(ByVal sheetName As String) As String()
    Dim servicesRng As Range
    Set servicesRng = SheetUtilities.getServiceHeaderRng(sheetName)
    ReDim services(servicesRng.Count - 1) As String
    Dim i As Long
    i = 1
    Do While i <= servicesRng.Count
        services(i - 1) = servicesRng.Cells.Item(1, i).value
        i = i + 1
    Loop
    
    loadServiceNames = services
End Function

Public Function getAddressRng(ByVal sheetName As String) As Range
    Set getAddressRng = getRng(sheetName, "A2", getServiceHeaderLastCell(sheetName))
End Function

Public Function getAddressVisitDataRng(ByVal sheetName As String) As Range
    Set getAddressVisitDataRng = Application.Union(getRng(sheetName, rxFirstCell, rxFirstCell), _
                                                   getRng(sheetName, serviceFirstCell, _
                                                          getServiceHeaderLastCell(sheetName)))
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

Public Sub CompareSheetCSV(ByVal Assert As Object, ByVal sheetName As String, ByVal csvPath As String, Optional ByVal rng As Range)
    Dim testArr() As String
    testArr = sheetToCSVArray(sheetName, rng)
    
    Dim correctArr() As String
    correctArr = getCSV(csvPath)
    
    Dim i As Long
    For i = LBound(correctArr, 1) To UBound(correctArr, 1)
        If i <= UBound(testArr) Then
            Assert.IsTrue StrComp(correctArr(i), testArr(i)) = 0, "Diff. at " & sheetName & " row " & i & " vs correct file: " & csvPath
        Else
            Assert.Fail "Diff. at " & sheetName & " row " & i & "vs correct file: " & csvPath
        End If
    Next i
End Sub

Public Sub ClearEmptyServices(ByVal sheetName As String)
    Dim servicesRng As Range
    Set servicesRng = getServiceHeaderRng(sheetName)
    
    Dim max As Long
    max = ActiveWorkbook.Worksheets.[_Default](sheetName).rows.Count
    
    Dim columnsToDelete As Range
    
    Dim i As Long
    i = 1
    Do While i <= servicesRng.Count
        If servicesRng.Cells.Item(max, i).End(xlUp).row = 1 Then
            Dim column As Range
            Set column = ActiveWorkbook.Worksheets.[_Default](sheetName).columns( _
                            i + SheetUtilities.firstServiceColumn - 1)
            If columnsToDelete Is Nothing Then
                Set columnsToDelete = column
            Else
                Set columnsToDelete = Application.Union(column, columnsToDelete)
            End If
        End If
        i = i + 1
    Loop
    
    If Not columnsToDelete Is Nothing Then columnsToDelete.EntireColumn.Delete
End Sub

Public Sub ClearSheet(ByVal sheetName As String)
    getAddressRng(sheetName).Clear
    getAddressVisitDataRng(sheetName).Clear
    getServiceHeaderRng(sheetName).Clear
End Sub

Public Sub ClearAll()
    getPastedRecordsRng.Clear
    getTotalsRng.value = 0
    getFinalReportRng.Clear
    
    Dim i As Long
    For i = 3 To ActiveWorkbook.Sheets.Count
        ClearSheet ActiveWorkbook.Sheets.[_Default](i).name
    Next
End Sub

Public Sub SortSheet(ByVal sheetName As String)
    Dim addressKey As String
    Select Case sheetName
        Case "Addresses", "Autocorrected"
            addressKey = "C2"
        Case "Needs Autocorrect", "Discards"
            addressKey = "F2"
    End Select
    
    getAddressRng(sheetName).Sort _
        key1:=ActiveWorkbook.Sheets.[_Default](sheetName).Range("B2"), _
        key2:=ActiveWorkbook.Sheets.[_Default](sheetName).Range(addressKey), _
        Order1:=xlDescending, Order2:=xlAscending, Header:=xlNo
End Sub

Public Sub SortAll()
    SortSheet "Addresses"
    SortSheet "Needs Autocorrect"
    SortSheet "Discards"
    SortSheet "Autocorrected"
End Sub

' Prints Collection
'@Ignore ParameterCanBeByVal
Public Sub PrintCollection(ByRef collectionResult As Collection)
    Dim i As Long
    Debug.Print ("[")
    For i = 1 To collectionResult.Count
        If TypeOf collectionResult.Item(i) Is Dictionary Then
            PrintJson collectionResult.Item(i)
        ElseIf TypeOf collectionResult.Item(i) Is Collection Then
            PrintCollection collectionResult.Item(i)
        Else
            Debug.Print """" & collectionResult.Item(i) & """"
        End If
        If i <> collectionResult.Count Then
            Debug.Print (",")
        End If
    Next
    Debug.Print ("]")
End Sub

' Prints JSON
'@Ignore ParameterCanBeByVal
Public Sub PrintJson(ByRef jsonResult As Scripting.Dictionary)
    Debug.Print ("{")
    
    Dim i As Long
    i = 0
    
    Do While i <= UBound(jsonResult.Keys)
        Dim key As Variant
        key = jsonResult.Keys(i)
        
        Debug.Print ("""" & key & """" & ":")
        If TypeOf jsonResult.Item(key) Is Collection Then
            PrintCollection jsonResult.Item(key)
        ElseIf TypeOf jsonResult.Item(key) Is Dictionary Then
            PrintJson jsonResult.Item(key)
        Else
            Debug.Print """" & (CStr(jsonResult.Item(key))) & """"
        End If
        
        If i < UBound(jsonResult.Keys) Then
            Debug.Print (",")
        End If
        
        i = i + 1
    Loop
    Debug.Print ("}")
End Sub


