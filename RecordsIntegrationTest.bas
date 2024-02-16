Attribute VB_Name = "RecordsIntegrationTest"
'@TestModule
'@Folder "City_Grant_Address_Report.test"


Option Explicit
Option Private Module

Private Assert As Object

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    ClearAll
    Autocorrect.printRemainingRequests 8000
End Sub

'@TestCleanup
Private Sub TestCleanup()
    ClearAll
    Autocorrect.printRemainingRequests 8000
End Sub

Private Sub PasteTestRecords(ByRef addressArr() As String)
    ActiveWorkbook.Worksheets.[_Default]("Interface").Activate
    
    ActiveSheet.Range("A9").Select
    
    Dim i As Long
    Dim fileArrLine() As String
    For i = 1 To UBound(addressArr, 1)
        If addressArr(i) <> vbNullString Then
            fileArrLine = Split(addressArr(i), ",")
            Dim j As Long
            For j = 0 To 11
                ActiveCell.Value = fileArrLine(j)
                ActiveCell.Offset(0, 1).Select
            Next j
            ActiveCell.Offset(1, -12).Select
        End If
    Next i
End Sub

'@TestMethod
Public Sub TestAllAddresses()
    On Error GoTo TestFail
    
    Dim testAddressesArr() As String
    testAddressesArr = getCSV(ThisWorkbook.path & "\testdata\test1addresses.csv")
    
    PasteTestRecords testAddressesArr
    
    addRecords
    
    CompareSheetCSV Assert, "Addresses", ActiveWorkbook.path & "\testdata\test1addresses_addressesoutput.csv"
    CompareSheetCSV Assert, "Interface", ActiveWorkbook.path & "\testdata\test1addresses_totalsoutput.csv", getTotalsRng
    CompareSheetCSV Assert, "Needs Autocorrect", ActiveWorkbook.path & "\testdata\test1addresses_autocorrectoutput.csv"
    CompareSheetCSV Assert, "Discards", ActiveWorkbook.path & "\testdata\test1addresses_discardsoutput.csv"
    CompareSheetCSV Assert, "Autocorrected", ActiveWorkbook.path & "\testdata\test1addresses_autocorrectedoutput.csv"
    
    Dim testExtraAddressesArr() As String
    testExtraAddressesArr = getCSV(ThisWorkbook.path & "\testdata\test2extraaddresses.csv")

    PasteTestRecords testExtraAddressesArr

    addRecords

    CompareSheetCSV Assert, "Addresses", ActiveWorkbook.path & "\testdata\test2extraaddresses_addressesoutput.csv"
    CompareSheetCSV Assert, "Interface", ActiveWorkbook.path & "\testdata\test2extraaddresses_totalsoutput.csv", getTotalsRng
    CompareSheetCSV Assert, "Needs Autocorrect", ActiveWorkbook.path & "\testdata\test2extraaddresses_autocorrectoutput.csv"
    CompareSheetCSV Assert, "Discards", ActiveWorkbook.path & "\testdata\test2extraaddresses_discardsoutput.csv"
    CompareSheetCSV Assert, "Autocorrected", ActiveWorkbook.path & "\testdata\test2extraaddresses_autocorrectedoutput.csv"

    Dim testAutocorrectAddressesArr() As String
    testAutocorrectAddressesArr = getCSV(ThisWorkbook.path & "\testdata\test3autocorrectaddresses.csv")

    PasteTestRecords testAutocorrectAddressesArr

    addRecords

    attemptValidation

    CompareSheetCSV Assert, "Addresses", ActiveWorkbook.path & "\testdata\test3autocorrectaddresses_addressesoutput.csv"
    CompareSheetCSV Assert, "Interface", ActiveWorkbook.path & "\testdata\test3autocorrectaddresses_totalsoutput.csv", getTotalsRng
    CompareSheetCSV Assert, "Needs Autocorrect", ActiveWorkbook.path & "\testdata\test3autocorrectaddresses_autocorrectoutput.csv"
    CompareSheetCSV Assert, "Discards", ActiveWorkbook.path & "\testdata\test3autocorrectaddresses_discardsoutput.csv"
    CompareSheetCSV Assert, "Autocorrected", ActiveWorkbook.path & "\testdata\test3autocorrectaddresses_autocorrectedoutput.csv"

    Assert.IsTrue Autocorrect.getRemainingRequests = 7980
'    Dim testMergeAutocorrectedAddressesArr() As String
'    testMergeAutocorrectedAddressesArr = getCSV(ThisWorkbook.path & "\testdata\testmergeautocorrectaddresses.csv")
'    PasteTestRecords testAutocorrectAddressesArr
'
'    addRecords
'
'    CompareSheetCSV assert, "Autocorrected", ActiveWorkbook.path & "\testdata\test3autocorrectaddresses_autocorrectedoutput.csv"
'
'    generateFinalReport
'
'    CompareSheetCSV assert, "Final Report", ActiveWorkbook.path & "\testdata\test3autocorrectaddresses_finalreportoutput.csv"
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub



'@TestMethod
Public Sub TestHandcorrected()
    ' TODO test against Diane corrected, get percentage correct
End Sub
