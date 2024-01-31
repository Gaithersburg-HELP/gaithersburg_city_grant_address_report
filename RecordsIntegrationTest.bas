Attribute VB_Name = "RecordsIntegrationTest"
'@TestModule
'@Folder "City_Grant_Address_Report.test"


Option Explicit
Option Private Module

Private assert As Object

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set assert = CreateObject("Rubberduck.AssertClass")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set assert = Nothing
    ClearAll
End Sub

'@TestMethod
Public Sub TestAllAddresses()
    On Error GoTo TestFail
    
    ClearAll
    
    Dim testAddressesArr() As String
    testAddressesArr = getCSV(ThisWorkbook.path & "\testdata\testaddresses.csv")
    
    ActiveWorkbook.Worksheets.[_Default]("Interface").Activate
    
    ActiveSheet.Range("A9").Select
    
    Dim i As Long
    Dim fileArrLine() As String
    For i = 1 To UBound(testAddressesArr, 1)
        If testAddressesArr(i) <> vbNullString Then
            fileArrLine = Split(testAddressesArr(i), ",")
            Dim j As Long
            For j = 0 To 11
                ActiveCell.Value = fileArrLine(j)
                ActiveCell.Offset(0, 1).Select
            Next j
            ActiveCell.Offset(1, -12).Select
        End If
    Next i
    
    addRecords
    
    CompareSheetCSV assert, "Addresses", ActiveWorkbook.path & "\testdata\testaddresses_addressesoutput.csv"
    CompareSheetCSV assert, "Interface", ActiveWorkbook.path & "\testdata\testaddresses_totalsoutput.csv", getTotalsRng
    CompareSheetCSV assert, "Needs Autocorrect", ActiveWorkbook.path & "\testdata\testaddresses_autocorrectoutput.csv"
    CompareSheetCSV assert, "Discards", ActiveWorkbook.path & "\testdata\testaddresses_discardsoutput.csv"
    CompareSheetCSV assert, "Autocorrected", ActiveWorkbook.path & "\testdata\testaddresses_autocorrectedoutput.csv"
    
    Exit Sub

TestFail:
    assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub TestLoadExtraAddressesAndWrite()
    ActiveWorkbook.Worksheets.[_Default]("Addresses").Activate
    ActiveSheet.Range("A2").Select
    
    ' TODO write some extra addresses
    
    ActiveWorkbook.Worksheets.[_Default]("Needs Autocorrect").Activate
    ActiveSheet.Range("A2").Select
    
    ' TODO write some extra addresses
    
    ActiveWorkbook.Worksheets.[_Default]("Discards").Activate
    ActiveSheet.Range("A2").Select
    
    ' TODO write some extra addresses
    
    ActiveWorkbook.Worksheets.[_Default]("Autocorrected").Activate
    ActiveSheet.Range("A2").Select
    
    ' TODO write some extra addresses
    
    ' assert new addresses are written after existing extra addresses
    ' assert existing extra addresses are merged with new addresses
    
End Sub

'@TestMethod
Public Sub TestAutocorrectAndFinalReport()
    ' TODO autocorrect method test
    generateFinalReport
    
    CompareSheetCSV assert, "Final Report", ActiveWorkbook.path & "\testdata\testaddresses__postvalidation_finalreportoutput.csv"
End Sub
