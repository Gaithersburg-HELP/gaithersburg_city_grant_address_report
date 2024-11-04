Attribute VB_Name = "SheetUtilities"
'@Folder("City_Grant_Address_Report.src")
Option Explicit

Public Const keyColumn As Long = 11
Public Const firstServiceColumn As Long = 19
Public Const mostRecentDateCell As String = "D1"

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

Public Function getVersionNum() As String
    getVersionNum = InterfaceSheet.Cells.Item(1, 1).value
End Function

Public Function getAPIKeyRng() As Range
    Set getAPIKeyRng = InterfaceSheet.Range("F1")
End Function

' Assumes services exist!
Public Function serviceFirstCell(ByVal sheetName As String) As String
    serviceFirstCell = ThisWorkbook.Worksheets.[_Default](sheetName) _
                         .Range("A1").Offset(0, firstServiceColumn - 1).address
End Function

Public Function rxFirstCell(ByVal sheetName As String) As String
    rxFirstCell = ThisWorkbook.Worksheets.[_Default](sheetName) _
                        .Range("A2").Offset(0, firstServiceColumn - 2).address
End Function

' Returns blank row after all data, assuming Column A is filled in last row
Public Function getBlankRow(ByVal sheetName As String) As Range
    Dim sheet As Worksheet
    Set sheet = ThisWorkbook.Worksheets.[_Default](sheetName)
    
    Set getBlankRow = sheet.rows.Item(sheet.rows.Item(sheet.rows.count).End(xlUp).row + 1)
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
        currentLastRow = sheet.Cells.Item(sheet.rows.count, i).End(xlUp).row
        If (currentLastRow > lastRow) Then lastRow = currentLastRow
        i = i + 1
    Loop
    
    Set getRng = sheet.Range(sheet.Range(firstCell), sheet.Cells.Item(lastRow, lastColNum))
End Function

Public Function getPastedRecordsRng() As Range
    Set getPastedRecordsRng = getRng(InterfaceSheet.Name, "A23", "O23")
End Function

Public Function getTotalsRng(ByVal totalService As TotalServiceType) As Range
    Select Case totalService '
        Case nonDelivery
            ' Include RxTotal for easy test comparison
            Set getTotalsRng = InterfaceSheet.Range("S3:V7")
        Case Delivery
            Set getTotalsRng = InterfaceSheet.Range("X3:AA6")
    End Select
End Function

Public Function getNonDeliveryTotalHeaderRng() As Range
    Set getNonDeliveryTotalHeaderRng = InterfaceSheet.Range("R1")
End Function

Public Function getDeliveryTotalHeaderRng() As Range
    Set getDeliveryTotalHeaderRng = InterfaceSheet.Range("W1")
End Function

Public Function getCountyRng() As Range
    Set getCountyRng = InterfaceSheet.Range("B9:CS20")
End Function

Public Function getNonRxReportRng() As Range
    Set getNonRxReportRng = getRng(NonRxReportSheet.Name, "A3", "P3")
End Function

Public Function getRxReportRng() As Range
    Set getRxReportRng = getRng(RxReportSheet.Name, "A3", "M3")
End Function

' Returns null if all services deleted
Private Function getServiceHeaderLastCell(ByVal sheetName As String) As String
    Dim lastCellAddr As String
    
    lastCellAddr = ThisWorkbook.Worksheets.[_Default](sheetName) _
                                      .Range("A1").Offset(0, firstServiceColumn - 2) _
                                      .End(xlToRight).address
    If lastCellAddr = "$XFD$1" Then
        getServiceHeaderLastCell = vbNullString
    Else
        getServiceHeaderLastCell = lastCellAddr
    End If
End Function

Public Function getServiceHeaderRng(ByVal sheetName As String) As Range
    Dim lastCellAddr As String
    lastCellAddr = getServiceHeaderLastCell(sheetName)
    
    If lastCellAddr = vbNullString Then
        Set getServiceHeaderRng = Nothing
    Else
        Set getServiceHeaderRng = ThisWorkbook.Worksheets.[_Default](sheetName) _
                                    .Range(serviceFirstCell(sheetName), lastCellAddr)
    End If
End Function

' Returns zero based service array
' If no services returns array with vbNullString at index 0
Public Function loadServiceNames(ByVal sheetName As String) As String()
    Dim servicesRng As Range
    Set servicesRng = SheetUtilities.getServiceHeaderRng(sheetName)
    
    If servicesRng Is Nothing Then
        Dim nullReturn(0) As String
        nullReturn(0) = vbNullString
        loadServiceNames = nullReturn
        Exit Function
    End If
    
    ReDim services(servicesRng.count - 1) As String
    Dim i As Long
    i = 1
    Do While i <= servicesRng.count
        services(i - 1) = servicesRng.Cells.Item(1, i).value
        i = i + 1
    Loop
    
    loadServiceNames = services
End Function

Public Function getMostRecentRng() As Range
    Set getMostRecentRng = InterfaceSheet.Range(mostRecentDateCell)
End Function

Public Function getAddressRng(ByVal sheetName As String) As Range
    Dim lastCellAddr As String
    lastCellAddr = getServiceHeaderLastCell(sheetName)
    
    If lastCellAddr = vbNullString Then
        lastCellAddr = rxFirstCell(sheetName)
    End If
    
    Set getAddressRng = getRng(sheetName, "A2", lastCellAddr)
End Function

Public Function getAddressVisitDataRng(ByVal sheetName As String) As Range
    Dim lastCellAddr As String
    lastCellAddr = getServiceHeaderLastCell(sheetName)
    
    If lastCellAddr = vbNullString Then
        Set getAddressVisitDataRng = getRng(sheetName, rxFirstCell(sheetName), rxFirstCell(sheetName))
    Else
        Set getAddressVisitDataRng = Application.Union(getRng(sheetName, rxFirstCell(sheetName), rxFirstCell(sheetName)), _
                                                   getRng(sheetName, serviceFirstCell(sheetName), _
                                                          lastCellAddr))
    End If
End Function

Public Function sheetToCSVArray(ByVal sheetName As String, Optional ByVal rng As Range = Nothing) As String()
    ' From https://stackoverflow.com/a/37038840/13342792
    
    If rng Is Nothing Then
        ThisWorkbook.Worksheets.[_Default](sheetName).Unprotect
        ThisWorkbook.Worksheets.[_Default](sheetName).UsedRange.Copy
    Else
        rng.Copy
    End If
    
    Dim TempWB As Workbook
    Set TempWB = Application.Workbooks.Add(1)
    ' NOTE following method sometimes fails (race condition?) if doing this in a new app
    ' or possibly started happening after protecting sheets or Application.Visible
    TempWB.Sheets.[_Default](1).Range("A1").PasteSpecial Paste:=xlPasteValues
    
    Dim MyFileName As String
    MyFileName = LibFileTools.GetLocalPath(ThisWorkbook.path) & "\test_" & sheetName & Format$(Time, "hh-mm-ss") & ".csv"
    
    Application.DisplayAlerts = False
    TempWB.SaveAs fileName:=MyFileName, FileFormat:=xlCSV, CreateBackup:=False, Local:=True
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
            Assert.isTrue StrComp(correctArr(i), testArr(i)) = 0, "Diff. at " & sheetName & " row " & i & " vs correct file: " & csvPath
        Else
            Assert.Fail "Diff. at " & sheetName & " row " & i & "vs correct file: " & csvPath
        End If
    Next i
End Sub

Public Sub ClearEmptyServices(ByVal sheetName As String)
    Dim servicesRng As Range
    Set servicesRng = getServiceHeaderRng(sheetName)
    
    If servicesRng Is Nothing Then Exit Sub
    
    Dim max As Long
    max = ThisWorkbook.Worksheets.[_Default](sheetName).rows.count
    
    Dim columnsToDelete As Range
    
    Dim i As Long
    i = 1
    Do While i <= servicesRng.count
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
    
    Dim serviceRng As Range
    Set serviceRng = getServiceHeaderRng(sheetName)
    If Not (serviceRng Is Nothing) Then serviceRng.Clear
End Sub

Public Sub ClearGburgTotals()
    Dim totalService As TotalServiceType
    For totalService = [_TotalServiceTypeFirst] To [_TotalServiceTypeLast]
        getTotalsRng(totalService).value = 0
    Next totalService
End Sub

Public Sub ClearAll()
    '@Ignore FunctionReturnValueDiscarded
    InterfaceButtons.MacroEntry ThisWorkbook.ActiveSheet
    Application.StatusBar = False
    
    getMostRecentRng.value = vbNullString
    getPastedRecordsRng.Clear
    InterfaceSheet.columns.Item("A").NumberFormat = "mm/dd/yyyy"
    
    ClearGburgTotals
    getNonDeliveryTotalHeaderRng.value = "Non-delivery"
    getDeliveryTotalHeaderRng.value = "Delivery"
    getCountyRng.value = 0
    getNonRxReportRng.Clear
    getRxReportRng.Clear
    
    Dim i As Long
    For i = 4 To ThisWorkbook.Sheets.count
        ClearSheet ThisWorkbook.Sheets.[_Default](i).Name
    Next
End Sub

Public Sub ClearAllPreserveDate()
    Dim mostRecentDate As String
    mostRecentDate = getMostRecentRng.value
    ClearAll
    getMostRecentRng.value = mostRecentDate
End Sub

Public Sub SortRange(ByVal rng As Range, ByVal sortOnValidFirst As Boolean)
    If rng.rows.count <= 1 Then
        Exit Sub
    End If
    
    Dim addressKey As String
    If sortOnValidFirst Then
        addressKey = "C1"
    Else
        addressKey = "F1"
    End If
    
    Dim row As Variant
    For Each row In rng.rows
        ' insert second word of address into temporary sort column to right of data
        ' NOTE don't use column to left of data, when tests fail then sometimes this
        ' column doesn't get deleted
        row.Offset(0, 1).Cells(1, rng.columns.count).value = LWordTrim(LWordTrim(row.Range(addressKey).value)(1))(0)
    Next row
    
    Dim rngWithSortCol As Range
    Set rngWithSortCol = rng.Resize(ColumnSize:=rng.columns.count + 1)

    '@Ignore ImplicitDefaultMemberAccess
    rngWithSortCol.Sort _
        key1:=rngWithSortCol.Cells.Item(1, rngWithSortCol.columns.count), _
        key2:=rng.Range(addressKey), _
        Order1:=xlAscending, Order2:=xlAscending, Header:=xlNo
        
    rngWithSortCol.columns.Item(rngWithSortCol.columns.count).EntireColumn.Clear
End Sub

Public Sub SortSheet(ByVal sheetName As String)
    Dim sortOnValidFirst As Boolean
    
    Select Case sheetName
        Case AddressesSheet.Name, AutocorrectedAddressesSheet.Name
            sortOnValidFirst = True
        ' Rubberduck Inspection bug
        '@Ignore UnreachableCase
        Case AutocorrectAddressesSheet.Name, DiscardsSheet.Name
            sortOnValidFirst = False
    End Select
    
    If sheetName = NonRxReportSheet.Name Or sheetName = RxReportSheet.Name Then
        ActiveWorkbook.Worksheets.[_Default](sheetName).Activate
        ActiveSheet.UsedRange.Offset(2, 0).Select
        
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
        SortRange getAddressRng(sheetName), sortOnValidFirst
    End If
End Sub

Public Function sortArr(ByRef arr() As Variant) As Variant()
    Dim sorted As Variant
    Set sorted = CreateObject("System.Collections.ArrayList")
    Dim element As Variant
    For Each element In arr
        sorted.Add element
    Next element
    sorted.Sort
    sortArr = sorted.toarray
End Function

Public Function cloneDict(ByVal dict As Scripting.Dictionary) As Scripting.Dictionary
    Dim newDict As Scripting.Dictionary
    Set newDict = CreateObject("Scripting.Dictionary")

    newDict.CompareMode = dict.CompareMode
    Dim key As Variant
    For Each key In dict.Keys
        newDict.Add key, dict.Item(key)
    Next

    Set cloneDict = newDict
End Function

Public Sub SortAll() ' TODO refactor? except for Final Report
    SortSheet AddressesSheet.Name
    SortSheet AutocorrectAddressesSheet.Name
    SortSheet DiscardsSheet.Name
    SortSheet AutocorrectedAddressesSheet.Name
End Sub

' Use JsonConverter.ConvertToJson instead of old PrintCollection and PrintJson

' Trims off first word including space afterwards
' Returns [trimmed first word, trimmed string (blank if only one word)]
Public Function LWordTrim(ByVal str As String) As String()
    Dim firstWord As String
    Dim spaceIndex As Long
    spaceIndex = InStr(1, str, " ", vbTextCompare)
    If (spaceIndex <> 0) Then
        firstWord = Left$(str, spaceIndex - 1)
        LWordTrim = Split(firstWord & "|" & Right$(str, Len(str) - spaceIndex), "|")
    Else
        LWordTrim = Split(str & "|", "|")
    End If
End Function


