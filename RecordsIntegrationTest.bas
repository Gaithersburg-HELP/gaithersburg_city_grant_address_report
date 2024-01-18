Attribute VB_Name = "RecordsIntegrationTest"
'@TestModule
'@Folder "City_Grant_Address_Report.test"


Option Explicit
Option Private Module

Private Assert As Object
Private Fakes As Object

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
    
    Dim testFileArray() As String
    testFileArray = getCSV(ThisWorkbook.path & "\testdata\testaddresses.csv")
    
    ActiveWorkbook.Worksheets.[_Default]("Interface").Activate
    
    ActiveSheet.Range("A9").Select
    
    Dim i As Long
    Dim fileArrLine() As String
    For i = 1 To UBound(testFileArray, 1)
        If testFileArray(i) <> vbNullString Then
            fileArrLine = Split(testFileArray(i), ",")
            Dim j As Long
            For j = 0 To 11
                ActiveCell.Value = fileArrLine(j)
                ActiveCell.Offset(0, 1).Select
            Next j
            ActiveCell.Offset(1, -12).Select
        End If
    Next i
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
    
    SheetUtilities.getPastedRecordsRng.Clear
    getTotalsRng.Value = 0
    getFinalReportRng.Clear
    getAddressesRng.Clear
    getDiscardsRng.Clear
    getAutocorrectRng.Clear
End Sub

Private Sub CompareSheetCSV(ByVal sheetName As String, ByVal csvPath As String, Optional ByVal rng As Range)
    Dim testArr() As String
    testArr = sheetToCSVArray(sheetName, rng)
    
    Dim correctArr() As String
    correctArr = getCSV(csvPath)
    
    Dim i As Long
    For i = LBound(correctArr, 1) To UBound(testArr, 1)
        Assert.IsTrue StrComp(correctArr(i), testArr(i)) = 0, "Difference at " & sheetName & " row " & i & ": " & correctArr(i) & "|" & testArr(i)
    Next i
End Sub


'@TestMethod
Public Sub TestAllAddresses()
    On Error GoTo TestFail
    
    addRecords
    
    CompareSheetCSV "Totals", ActiveWorkbook.path & "\testdata\testaddresses_totalsoutput.csv", getTotalsRng
    CompareSheetCSV "Addresses", ActiveWorkbook.path & "\testdata\testaddresses_addressesoutput.csv"
    CompareSheetCSV "Invalid Discards", ActiveWorkbook.path & "\testdata\testaddresses_discardsoutput.csv"
    CompareSheetCSV "Autocorrected Addresses", ActiveWorkbook.path & "\testdata\testaddresses_autocorrectoutput.csv"

    generateFinalReport
    
    CompareSheetCSV "Final Report", ActiveWorkbook.path & "\testdata\testaddresses_finalreportoutput.csv"
    
    Exit Sub

TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
