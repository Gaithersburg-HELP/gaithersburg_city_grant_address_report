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
    
    ClearAll
    
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
    Set assert = Nothing
    
    ClearAll
End Sub

'@TestMethod
Public Sub TestAllAddresses()
    On Error GoTo TestFail
    
    addRecords
    
    CompareSheetCSV assert, "Addresses", ActiveWorkbook.path & "\testdata\testaddresses_addressesoutput.csv"
    CompareSheetCSV assert, "Totals", ActiveWorkbook.path & "\testdata\testaddresses_totalsoutput.csv", getTotalsRng
    CompareSheetCSV assert, "Invalid Discards", ActiveWorkbook.path & "\testdata\testaddresses_discardsoutput.csv"
    CompareSheetCSV assert, "Autocorrected Addresses", ActiveWorkbook.path & "\testdata\testaddresses_autocorrectoutput.csv"

    generateFinalReport
    
    CompareSheetCSV assert, "Final Report", ActiveWorkbook.path & "\testdata\testaddresses_finalreportoutput.csv"
    
    Exit Sub

TestFail:
    assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub TestLoadAddressesAndAutocorrect()
    'TODO test
End Sub
