Attribute VB_Name = "SheetUtilities"
'@Folder("City_Grant_Address_Report.src")
Option Explicit

Public Const firstServiceColumn As Long = 19

Public Enum CountyTotalCols
    countymonth = 1
    householdDuplicate = 2
    householdUnduplicate = 3
    individualDuplicate = 4
    individualUnduplicate = 5
    childrenDuplicate = 6
    adultDuplicate = 7
    poundsFood = 8
    
    zip20906AshtonAspenHill = 10
    zip20906SilverSpring = 84
    
    zip20916AshtonAspenHill = 11
    zip20916SilverSpring = 93
    
    zip20815Bethesda = 16
    zip20815ChevyChaseClarksburg = 27
    
    zip20825Bethesda = 20
    zip20825ChevyChaseClarksburg = 28
    
    zip20852Bethesda = 22
    zip20852Rockville = 70
    
    zip20904ColesvilleDamascus = 30
    zip20904SilverSpring = 82
    
    zip20905ColesvilleDamascus = 31
    zip20905SilverSpring = 83
    
    zip20914ColesvilleDamascus = 32
    zip20914SilverSpring = 91
    
    zip20874DarnestownDerwoodDickerson = 34
    zip20874GarrettParkGermantownGlenEcho = 48
    
    zip20878DarnestownDerwoodDickerson = 35
    zip20878Gaithersburg = 39
    zip20878PoolesvillePotomac = 64
    
    zip20855DarnestownDerwoodDickerson = 36
    zip20855Rockville = 73
    
    zip20877Gaithersburg = 38
    zip20877MontgomeryVillageOlney = 56
    
    zip20882Gaithersburg = 41
    zip20882KensingtonLaytonsville = 55
    
    zip20886Gaithersburg = 45
    zip20886MontgomeryVillageOlney = 58
    
    zip20879Gaithersburg = 40
    zip20879KensingtonLaytonsville = 54
    zip20879MontgomeryVillageOlney = 57
     
    zip20854PoolesvillePotomac = 62
    zip20854Rockville = 72
    
    zip20859PoolesvillePotomac = 63
    zip20859Rockville = 74
    
    zip20912SandySpringSpencervilleTakomaPark = 77
    zip20912SilverSpring = 89
    
    zip20913SandySpringSpencervilleTakomaPark = 78
    zip20913SilverSpring = 90
    
    zip20902SilverSpring = 80
    zip20902WashingtonGroveWheaton = 96
    
    zip20915SilverSpring = 92
    zip20915WashingtonGroveWheaton = 97
End Enum

Public Function uniqueCountyZipCols() As Scripting.Dictionary
    Dim cols As Scripting.Dictionary
    Set cols = New Scripting.Dictionary
        cols.Add "20861", 9
    cols.Add "20839", 12
    cols.Add "20838", 13
    cols.Add "20813", 14
    cols.Add "20814", 15
    cols.Add "20816", 17
    cols.Add "20817", 18
    cols.Add "20824", 19
    cols.Add "20827", 21
    cols.Add "20841", 23
    cols.Add "20862", 24
    cols.Add "20866", 25
    cols.Add "20818", 26
    cols.Add "20871", 29
    cols.Add "20872", 33
    cols.Add "20842", 37
    cols.Add "20883", 42
    cols.Add "20884", 43
    cols.Add "20885", 44
    cols.Add "20898", 46
    cols.Add "20896", 47
    cols.Add "20875", 49
    cols.Add "20876", 50
    cols.Add "20812", 51
    cols.Add "20891", 52
    cols.Add "20895", 53
    cols.Add "20830", 59
    cols.Add "20832", 60
    cols.Add "20837", 61
    cols.Add "20847", 65
    cols.Add "20848", 66
    cols.Add "20849", 67
    cols.Add "20850", 68
    cols.Add "20851", 69
    cols.Add "20853", 71
    cols.Add "20860", 75
    cols.Add "20868", 76
    cols.Add "20901", 79
    cols.Add "20903", 81
    cols.Add "20907", 85
    cols.Add "20908", 86
    cols.Add "20910", 87
    cols.Add "20911", 88
    cols.Add "20918", 94
    cols.Add "20880", 95
    Set uniqueCountyZipCols = cols
End Function


Public Function serviceFirstCell() As String
    serviceFirstCell = ActiveSheet.Range("A1").offset(0, firstServiceColumn - 1).address
End Function

Public Function rxFirstCell() As String
    rxFirstCell = ActiveSheet.Range("A2").offset(0, firstServiceColumn - 2).address
End Function

' Returns blank row after all data, assuming Column A is filled in last row
Public Function getBlankRow(ByVal sheetName As String) As Range
    Dim sheet As Worksheet
    Set sheet = ThisWorkbook.Worksheets.[_Default](sheetName)
    
    Set getBlankRow = sheet.rows.Item(sheet.rows.Item(sheet.rows.Count).End(xlUp).row + 1)
End Function

' Returns all data below (all cells between firstCell and lastCol) including blanks and firstCell
Public Function getRng(ByVal sheetName As String, ByVal firstCell As String, ByVal lastCol As String) As Range
    Dim sheet As Worksheet
    Set sheet = ThisWorkbook.Worksheets.[_Default](sheetName)
        
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
    Set getPastedRecordsRng = getRng("Interface", "A23", "O23")
End Function

Public Function getTotalsRng() As Range
    Set getTotalsRng = ThisWorkbook.Worksheets.[_Default]("Interface").Range("N2:Q6")
End Function

Public Function getCountyRng() As Range
    Set getCountyRng = ThisWorkbook.Worksheets.[_Default]("Interface").Range("B9:CS20")
End Function

Public Function getFinalReportRng() As Range
    Set getFinalReportRng = getRng("Final Report", "A2", "M2")
End Function

Private Function getServiceHeaderLastCell(ByVal sheetName As String) As String
    getServiceHeaderLastCell = ThisWorkbook.Worksheets.[_Default](sheetName) _
                                      .Range("A1").offset(0, firstServiceColumn - 2) _
                                      .End(xlToRight).address
End Function

Public Function getServiceHeaderRng(ByVal sheetName As String) As Range
    Set getServiceHeaderRng = ThisWorkbook.Worksheets.[_Default](sheetName) _
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
     
    Set CurrentWB = ThisWorkbook
    
    If rng Is Nothing Then
        ThisWorkbook.Worksheets.[_Default](sheetName).UsedRange.Copy
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
    max = ThisWorkbook.Worksheets.[_Default](sheetName).rows.Count
    
    Dim columnsToDelete As Range
    
    Dim i As Long
    i = 1
    Do While i <= servicesRng.Count
        If servicesRng.Cells.Item(max, i).End(xlUp).row = 1 Then
            Dim column As Range
            Set column = ThisWorkbook.Worksheets.[_Default](sheetName).columns( _
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
    getCountyRng.value = 0
    getFinalReportRng.Clear
    
    Dim i As Long
    For i = 3 To ThisWorkbook.Sheets.Count
        ClearSheet ThisWorkbook.Sheets.[_Default](i).Name
    Next
End Sub

Public Sub DisableAllFilters()
    Dim i As Long
    For i = 1 To ThisWorkbook.Sheets.Count
        ThisWorkbook.Sheets.[_Default](i).AutoFilterMode = False
    Next
End Sub

Public Sub SortSheet(ByVal sheetName As String)
    Dim addressKey As String
    Select Case sheetName
    Case "Addresses", "Autocorrected", "Final Report"
        addressKey = "C2"
    Case "Needs Autocorrect", "Discards"
        addressKey = "F2"
    End Select
    
    If sheetName = "Final Report" Then
        ThisWorkbook.Sheets.[_Default]("Final Report").Select
        ThisWorkbook.Sheets.[_Default]("Final Report").Range("A2:O2").Select
        ActiveSheet.Range(selection, selection.End(xlDown)).Select
        
        With ActiveSheet.Sort
            .SortFields.Clear
            .SortFields.Add key:=selection.columns(3), Order:=xlAscending
            .SortFields.Add key:=selection.columns(2), Order:=xlAscending
            .SortFields.Add key:=selection.columns(4), Order:=xlAscending
            .SortFields.Add key:=selection.columns(6), Order:=xlAscending
            .Header = xlNo
            .SetRange selection
            .Apply
        End With
    Else
        getAddressRng(sheetName).Sort _
        key1:=ThisWorkbook.Sheets.[_Default](sheetName).Range("B2"), _
        key2:=ThisWorkbook.Sheets.[_Default](sheetName).Range(addressKey), _
        Order1:=xlDescending, Order2:=xlAscending, Header:=xlNo
    End If
End Sub

Public Sub SortAll() ' TODO refactor? except for Final Report
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


