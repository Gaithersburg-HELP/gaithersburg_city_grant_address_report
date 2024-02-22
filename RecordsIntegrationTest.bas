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
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    ClearAll
    autocorrect.printRemainingRequests 8000
End Sub

'@TestCleanup
Private Sub TestCleanup()
    ClearAll
    autocorrect.printRemainingRequests 8000
End Sub

Private Sub PasteTestRecords(ByRef addressArr() As String)
    ActiveWorkbook.Worksheets.[_Default]("Interface").Select
    getPastedRecordsRng.Cells.Item(1, 1).Select
    
    Dim i As Long
    Dim fileArrLine() As String
    For i = 1 To UBound(addressArr, 1)
        If addressArr(i) <> vbNullString Then
            fileArrLine = Split(addressArr(i), ",")
            Dim j As Long
            For j = 0 To 12 ' TODO when adding adult/child. Not UBound because of test notes
                ActiveCell.value = fileArrLine(j)
                ActiveCell.offset(0, 1).Select
            Next j
            ActiveSheet.Cells(ActiveCell.row + 1, 1).Select
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

    Assert.IsTrue autocorrect.getRemainingRequests = 7980


    Dim testMergeAutocorrectedAddressesArr() As String
    testMergeAutocorrectedAddressesArr = getCSV(ThisWorkbook.path & "\testdata\test4mergeaddresses.csv")
    PasteTestRecords testMergeAutocorrectedAddressesArr

    addRecords

    CompareSheetCSV Assert, "Addresses", ActiveWorkbook.path & "\testdata\test4mergeaddresses_addressesoutput.csv"
    CompareSheetCSV Assert, "Interface", ActiveWorkbook.path & "\testdata\test4mergeaddresses_totalsoutput.csv", getTotalsRng
    CompareSheetCSV Assert, "Discards", ActiveWorkbook.path & "\testdata\test4mergeaddresses_discardsoutput.csv"
    CompareSheetCSV Assert, "Autocorrected", ActiveWorkbook.path & "\testdata\test4mergeaddresses_autocorrectedoutput.csv"

    Fakes.MsgBox.Returns vbYes

    InterfaceButtons.confirmDiscardAll

    ActiveWorkbook.Worksheets.[_Default]("Discards").Select
    Union(ActiveWorkbook.Worksheets.[_Default]("Discards").Range("A3:A6"), _
          ActiveWorkbook.Worksheets.[_Default]("Discards").Range("A9:A12"), _
          ActiveWorkbook.Worksheets.[_Default]("Discards").Range("A14:A15")).Select
    InterfaceButtons.confirmRestoreSelectedDiscard

    ActiveWorkbook.Worksheets.[_Default]("Needs Autocorrect").Select
    Union(ActiveWorkbook.Worksheets.[_Default]("Needs Autocorrect").Range("A6"), _
          ActiveWorkbook.Worksheets.[_Default]("Needs Autocorrect").Range("A9")).Select
    InterfaceButtons.confirmDiscardSelected

    ActiveWorkbook.Worksheets.[_Default]("Addresses").Select
    Union(ActiveWorkbook.Worksheets.[_Default]("Addresses").Range("A3"), _
          ActiveWorkbook.Worksheets.[_Default]("Addresses").Range("A8")).Select
    InterfaceButtons.confirmMoveAutocorrect

    ActiveWorkbook.Worksheets.[_Default]("Needs Autocorrect").Select
    ActiveWorkbook.Worksheets.[_Default]("Needs Autocorrect").Range("D2").value = "Ste 102"
    ActiveWorkbook.Worksheets.[_Default]("Needs Autocorrect").Range("C3").value = "13-15 E Deer Park Dr"
    ActiveWorkbook.Worksheets.[_Default]("Needs Autocorrect").Range("D3").value = "Ste 202"
    ActiveWorkbook.Worksheets.[_Default]("Needs Autocorrect").Range("D4").value = "Unit 102"
    ActiveWorkbook.Worksheets.[_Default]("Needs Autocorrect").Range("D5").value = "Unit 102"
    ActiveWorkbook.Worksheets.[_Default]("Needs Autocorrect").Range("D8").value = "Apt T3"
    ActiveWorkbook.Worksheets.[_Default]("Needs Autocorrect").Range("D9").value = "Apt 103"
    ActiveWorkbook.Worksheets.[_Default]("Needs Autocorrect").Range("D10").value = "Ste 100"
    ActiveWorkbook.Worksheets.[_Default]("Needs Autocorrect").Range("D11").value = "Apt 1"

    Union(ActiveWorkbook.Worksheets.[_Default]("Needs Autocorrect").Range("A2:A5"), _
          ActiveWorkbook.Worksheets.[_Default]("Needs Autocorrect").Range("A8:A11")).Select
    InterfaceButtons.toggleUserVerified

    CompareSheetCSV Assert, "Addresses", ActiveWorkbook.path & "\testdata\test5usereditsaddresses_addressesoutput.csv"
    CompareSheetCSV Assert, "Interface", ActiveWorkbook.path & "\testdata\test5usereditsaddresses_totalsoutput.csv", getTotalsRng
    CompareSheetCSV Assert, "Needs Autocorrect", ActiveWorkbook.path & "\testdata\test5usereditsaddresses_autocorrectoutput.csv"
    CompareSheetCSV Assert, "Discards", ActiveWorkbook.path & "\testdata\test5usereditsaddresses_discardsoutput.csv"
    CompareSheetCSV Assert, "Autocorrected", ActiveWorkbook.path & "\testdata\test5usereditsaddresses_autocorrectedoutput.csv"

    InterfaceButtons.confirmAttemptValidation
    InterfaceButtons.confirmGenerateFinalReport
'
'    ' TODO TEST Odend'hal, O'neill
'    CompareSheetCSV Assert, "Addresses", ActiveWorkbook.path & "\testdata\test6validateduseredits_addressesoutput.csv"
'    CompareSheetCSV Assert, "Interface", ActiveWorkbook.path & "\testdata\test6validateduseredits_totalsoutput.csv", getTotalsRng
'    CompareSheetCSV Assert, "Autocorrected", ActiveWorkbook.path & "\testdata\test6validateduseredits_autocorrectedoutput.csv"
'    CompareSheetCSV Assert, "Final Report", ActiveWorkbook.path & "\testdata\test6validateduseredits_finalreportoutput.csv"


'
'    ' TODO test delete service column, generate final report
'    InterfaceButtons.confirmDeleteService
'    InterfaceButtons.confirmGenerateFinalReport
'
'    ' TODO move one record to needs autocorrect, test delete all visit data
'    InterfaceButtons.confirmDeleteAllVisitData
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub



'@TestMethod
Public Sub TestHandcorrected()
    ' TODO test against Diane corrected, get percentage correct
End Sub
