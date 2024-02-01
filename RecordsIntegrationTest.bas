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
    
    ClearAll
    
    Dim testAddressesArr() As String
    testAddressesArr = getCSV(ThisWorkbook.path & "\testdata\test1addresses.csv")
    
    PasteTestRecords testAddressesArr
    
    addRecords
    
    CompareSheetCSV assert, "Addresses", ActiveWorkbook.path & "\testdata\test1addresses_addressesoutput.csv"
    CompareSheetCSV assert, "Interface", ActiveWorkbook.path & "\testdata\test1addresses_totalsoutput.csv", getTotalsRng
    CompareSheetCSV assert, "Needs Autocorrect", ActiveWorkbook.path & "\testdata\test1addresses_autocorrectoutput.csv"
    CompareSheetCSV assert, "Discards", ActiveWorkbook.path & "\testdata\test1addresses_discardsoutput.csv"
    CompareSheetCSV assert, "Autocorrected", ActiveWorkbook.path & "\testdata\test1addresses_autocorrectedoutput.csv"
    
    Dim testAddAddressesArr() As String
    testAddAddressesArr = getCSV(ThisWorkbook.path & "\testdata\testaddaddresses.csv")
    
    PasteTestRecords testAddAddressesArr
    
    addRecords
    
    CompareSheetCSV assert, "Addresses", ActiveWorkbook.path & "\testdata\test2extraaddresses_addressesoutput.csv"
    CompareSheetCSV assert, "Interface", ActiveWorkbook.path & "\testdata\test2extraaddresses_totalsoutput.csv", getTotalsRng
    CompareSheetCSV assert, "Needs Autocorrect", ActiveWorkbook.path & "\testdata\test2extraaddresses_autocorrectoutput.csv"
    CompareSheetCSV assert, "Discards", ActiveWorkbook.path & "\testdata\test2extraaddresses_discardsoutput.csv"
    CompareSheetCSV assert, "Autocorrected", ActiveWorkbook.path & "\testdata\test2extraaddresses_autocorrectedoutput.csv"
    
    Dim testAutocorrectAddressesArr() As String
    testAutocorrectAddressesArr = getCSV(ThisWorkbook.path & "\testdata\testautocorrectaddresses.csv")
    
    PasteTestRecords testAutocorrectAddressesArr
    
    addRecords
    
    attemptValidation
    
    generateFinalReport
    
    CompareSheetCSV assert, "Addresses", ActiveWorkbook.path & "\testdata\test3autocorrectaddresses_addressesoutput.csv"
    CompareSheetCSV assert, "Interface", ActiveWorkbook.path & "\testdata\test3autocorrectaddresses_totalsoutput.csv", getTotalsRng
    CompareSheetCSV assert, "Needs Autocorrect", ActiveWorkbook.path & "\testdata\test3autocorrectaddresses_autocorrectoutput.csv"
    CompareSheetCSV assert, "Discards", ActiveWorkbook.path & "\testdata\test3autocorrectaddresses_discardsoutput.csv"
    CompareSheetCSV assert, "Autocorrected", ActiveWorkbook.path & "\testdata\test3autocorrectaddresses_autocorrectedoutput.csv"
    CompareSheetCSV assert, "Final Report", ActiveWorkbook.path & "\testdata\test3autocorrectaddresses_finalreportoutput.csv"
    Exit Sub

TestFail:
    assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub



'@TestMethod
Public Sub TestHandcorrected()
    ' TODO test against Diane corrected, get percentage correct
End Sub
