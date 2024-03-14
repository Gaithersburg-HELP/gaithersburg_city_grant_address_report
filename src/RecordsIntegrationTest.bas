Attribute VB_Name = "RecordsIntegrationTest"
'@IgnoreModule FunctionReturnValueDiscarded
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

    Application.Visible = False
    Application.ScreenUpdating = False
    MacroEntry InterfaceSheet
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
    
    MacroExit InterfaceSheet
    Application.ScreenUpdating = True
    Application.Visible = True
End Sub

'@TestInitialize
Private Sub TestInitialize()
    ClearAll
    autocorrect.printRemainingRequests 8000
End Sub

'@TestCleanup
Private Sub TestCleanup()
    MacroEntry InterfaceSheet
    ClearAll
    autocorrect.printRemainingRequests 8000
End Sub

Private Sub PasteTestRecords(ByRef addressArr() As String)
    InterfaceSheet.Select
    getPastedRecordsRng.Cells.Item(1, 1).Select
    
    Dim i As Long
    Dim fileArrLine() As String
    For i = 1 To UBound(addressArr, 1)
        If addressArr(i) <> vbNullString Then
            fileArrLine = Split(addressArr(i), ",")
            Dim j As Long
            For j = 0 To 14 ' TODO when adding adult/child. Not UBound because of test notes
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
    
    CompareSheetCSV Assert, "Addresses", ThisWorkbook.path & "\testdata\test1addresses_addressesoutput.csv"
    CompareSheetCSV Assert, "Interface", ThisWorkbook.path & "\testdata\test1addresses_totalsoutput.csv", getTotalsRng
    CompareSheetCSV Assert, "Interface", ThisWorkbook.path & "\testdata\test1addresses_countytotalsoutput.csv", getCountyRng
    CompareSheetCSV Assert, "Needs Autocorrect", ThisWorkbook.path & "\testdata\test1addresses_autocorrectoutput.csv"
    CompareSheetCSV Assert, "Discards", ThisWorkbook.path & "\testdata\test1addresses_discardsoutput.csv"
    CompareSheetCSV Assert, "Autocorrected", ThisWorkbook.path & "\testdata\test1addresses_autocorrectedoutput.csv"
    
    Dim testExtraAddressesArr() As String
    testExtraAddressesArr = getCSV(ThisWorkbook.path & "\testdata\test2extraaddresses.csv")

    PasteTestRecords testExtraAddressesArr

    addRecords

    CompareSheetCSV Assert, "Addresses", ThisWorkbook.path & "\testdata\test2extraaddresses_addressesoutput.csv"
    CompareSheetCSV Assert, "Interface", ThisWorkbook.path & "\testdata\test2extraaddresses_totalsoutput.csv", getTotalsRng
    CompareSheetCSV Assert, "Needs Autocorrect", ThisWorkbook.path & "\testdata\test2extraaddresses_autocorrectoutput.csv"
    CompareSheetCSV Assert, "Discards", ThisWorkbook.path & "\testdata\test2extraaddresses_discardsoutput.csv"
    CompareSheetCSV Assert, "Autocorrected", ThisWorkbook.path & "\testdata\test2extraaddresses_autocorrectedoutput.csv"

    Dim testAutocorrectAddressesArr() As String
    testAutocorrectAddressesArr = getCSV(ThisWorkbook.path & "\testdata\test3autocorrectaddresses.csv")

    PasteTestRecords testAutocorrectAddressesArr

    addRecords

    attemptValidation

    CompareSheetCSV Assert, "Addresses", ThisWorkbook.path & "\testdata\test3autocorrectaddresses_addressesoutput.csv"
    CompareSheetCSV Assert, "Interface", ThisWorkbook.path & "\testdata\test3autocorrectaddresses_totalsoutput.csv", getTotalsRng
    CompareSheetCSV Assert, "Needs Autocorrect", ThisWorkbook.path & "\testdata\test3autocorrectaddresses_autocorrectoutput.csv"
    CompareSheetCSV Assert, "Discards", ThisWorkbook.path & "\testdata\test3autocorrectaddresses_discardsoutput.csv"
    CompareSheetCSV Assert, "Autocorrected", ThisWorkbook.path & "\testdata\test3autocorrectaddresses_autocorrectedoutput.csv"

    Assert.IsTrue autocorrect.getRemainingRequests = 7980


    Dim testMergeAutocorrectedAddressesArr() As String
    testMergeAutocorrectedAddressesArr = getCSV(ThisWorkbook.path & "\testdata\test4mergeaddresses.csv")
    PasteTestRecords testMergeAutocorrectedAddressesArr

    addRecords

    CompareSheetCSV Assert, "Addresses", ThisWorkbook.path & "\testdata\test4mergeaddresses_addressesoutput.csv"
    CompareSheetCSV Assert, "Interface", ThisWorkbook.path & "\testdata\test4mergeaddresses_totalsoutput.csv", getTotalsRng
    CompareSheetCSV Assert, "Discards", ThisWorkbook.path & "\testdata\test4mergeaddresses_discardsoutput.csv"
    CompareSheetCSV Assert, "Autocorrected", ThisWorkbook.path & "\testdata\test4mergeaddresses_autocorrectedoutput.csv"

    Fakes.MsgBox.Returns vbYes

    InterfaceButtons.confirmDiscardAll
    MacroEntry InterfaceSheet

    DiscardsSheet.Select
    Union(DiscardsSheet.Range("A2:A6"), _
          DiscardsSheet.Range("A9:A13")). _
          Select
    InterfaceButtons.confirmRestoreSelectedDiscard
    MacroEntry InterfaceSheet
    
    AutocorrectAddressesSheet.Select
    Union(AutocorrectAddressesSheet.Range("A4"), _
          AutocorrectAddressesSheet.Range("A7:A8")).Select
    InterfaceButtons.confirmDiscardSelected
    MacroEntry InterfaceSheet
    
    AddressesSheet.Select
    Union(AddressesSheet.Range("A3"), _
          AddressesSheet.Range("A7"), _
          AddressesSheet.Range("A14")).Select
    InterfaceButtons.confirmMoveAutocorrect
    MacroEntry InterfaceSheet
    
    AutocorrectAddressesSheet.Select

    AutocorrectAddressesSheet.Range("C5").value = "13-15 E Deer Park Dr"
    AutocorrectAddressesSheet.Range("D5").value = "Ste 102"
    AutocorrectAddressesSheet.Range("C6").value = "13-15 E Deer Park Dr"
    AutocorrectAddressesSheet.Range("D6").value = "Ste 202"
    AutocorrectAddressesSheet.Range("D2").value = "Unit 102"
    AutocorrectAddressesSheet.Range("D3").value = "Unit 102"
    AutocorrectAddressesSheet.Range("B7").value = True
    AutocorrectAddressesSheet.Range("D8").value = "Apt 103"
    AutocorrectAddressesSheet.Range("D9").value = "Ste 100"
    AutocorrectAddressesSheet.Range("D11").value = "Apt 1"

    Union(AutocorrectAddressesSheet.Range("A2:A6"), AutocorrectAddressesSheet.Range("A8:A9"), _
          AutocorrectAddressesSheet.Range("A11")).Select
    InterfaceButtons.toggleUserVerified
    MacroEntry InterfaceSheet

    AutocorrectedAddressesSheet.Select
    Union(AutocorrectedAddressesSheet.Range("A3"), _
          AutocorrectedAddressesSheet.Range("A5")).Select
    InterfaceButtons.toggleUserVerifiedAutocorrected
    MacroEntry InterfaceSheet

    CompareSheetCSV Assert, "Addresses", ThisWorkbook.path & "\testdata\test5usereditsaddresses_addressesoutput.csv"
    CompareSheetCSV Assert, "Interface", ThisWorkbook.path & "\testdata\test5usereditsaddresses_totalsoutput.csv", getTotalsRng
    CompareSheetCSV Assert, "Needs Autocorrect", ThisWorkbook.path & "\testdata\test5usereditsaddresses_autocorrectoutput.csv"
    CompareSheetCSV Assert, "Discards", ThisWorkbook.path & "\testdata\test5usereditsaddresses_discardsoutput.csv"
    CompareSheetCSV Assert, "Autocorrected", ThisWorkbook.path & "\testdata\test5usereditsaddresses_autocorrectedoutput.csv"

    InterfaceButtons.confirmAttemptValidation
    MacroEntry InterfaceSheet
    
    InterfaceButtons.confirmGenerateFinalReport
    MacroEntry InterfaceSheet

    CompareSheetCSV Assert, "Addresses", ThisWorkbook.path & "\testdata\test6validateduseredits_addressesoutput.csv"
    CompareSheetCSV Assert, "Interface", ThisWorkbook.path & "\testdata\test6validateduseredits_totalsoutput.csv", getTotalsRng
    CompareSheetCSV Assert, "Interface", ThisWorkbook.path & "\testdata\test6validateduseredits_countyoutput.csv", getCountyRng
    CompareSheetCSV Assert, "Autocorrected", ThisWorkbook.path & "\testdata\test6validateduseredits_autocorrectedoutput.csv"
    CompareSheetCSV Assert, "Final Report", ThisWorkbook.path & "\testdata\test6validateduseredits_finalreportoutput.csv"

    AddressesSheet.Select
    AddressesSheet.Range("A6").Select
    InterfaceButtons.confirmMoveAutocorrect
    MacroEntry InterfaceSheet
    
    InterfaceButtons.confirmDeleteAllVisitData
    MacroEntry InterfaceSheet

    CompareSheetCSV Assert, "Addresses", ThisWorkbook.path & "\testdata\test7deletedata_addressesoutput.csv"
    CompareSheetCSV Assert, "Interface", ThisWorkbook.path & "\testdata\test7deletedata_totalsoutput.csv", getTotalsRng
    CompareSheetCSV Assert, "Interface", ThisWorkbook.path & "\testdata\test7deletedata_countyoutput.csv", getCountyRng
    CompareSheetCSV Assert, "Needs Autocorrect", ThisWorkbook.path & "\testdata\test7deletedata_autocorrectoutput.csv"
    CompareSheetCSV Assert, "Discards", ThisWorkbook.path & "\testdata\test7deletedata_discardsoutput.csv"
    CompareSheetCSV Assert, "Autocorrected", ThisWorkbook.path & "\testdata\test7deletedata_autocorrectedoutput.csv"
    CompareSheetCSV Assert, "Final Report", ThisWorkbook.path & "\testdata\test7deletedata_finalreportoutput.csv"

    AddressesSheet.Select
    AddressesSheet.Range("A2").Select
    InterfaceButtons.confirmMoveAutocorrect

    AutocorrectAddressesSheet.Select
    AutocorrectAddressesSheet.Range("A2").Select
    InterfaceButtons.toggleUserVerified
    InterfaceButtons.confirmDiscardSelected
    InterfaceButtons.confirmDiscardAll

    DiscardsSheet.Select
    DiscardsSheet.Range("A2").Select
    InterfaceButtons.confirmRestoreSelectedDiscard

    AutocorrectedAddressesSheet.Select
    AutocorrectedAddressesSheet.Range("A2").Select
    InterfaceButtons.toggleUserVerifiedAutocorrected

    InterfaceSheet.Select
    PasteTestRecords testMergeAutocorrectedAddressesArr

    InterfaceButtons.confirmAddRecords
    InterfaceButtons.confirmAttemptValidation
    InterfaceButtons.confirmGenerateFinalReport

    CompareSheetCSV Assert, "Addresses", ThisWorkbook.path & "\testdata\test8noservices_addressesoutput.csv"
    CompareSheetCSV Assert, "Interface", ThisWorkbook.path & "\testdata\test8noservices_totalsoutput.csv", getTotalsRng
    CompareSheetCSV Assert, "Interface", ThisWorkbook.path & "\testdata\test8noservices_countyoutput.csv", getCountyRng
    CompareSheetCSV Assert, "Needs Autocorrect", ThisWorkbook.path & "\testdata\test8noservices_autocorrectoutput.csv"
    CompareSheetCSV Assert, "Discards", ThisWorkbook.path & "\testdata\test8noservices_discardsoutput.csv"
    CompareSheetCSV Assert, "Autocorrected", ThisWorkbook.path & "\testdata\test8noservices_autocorrectedoutput.csv"
    CompareSheetCSV Assert, "Final Report", ThisWorkbook.path & "\testdata\test8noservices_finalreportoutput.csv"
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestToggleUserVerifiedAutocorrect()
    On Error GoTo TestFail
    
    Dim testAddressesArr() As String
    testAddressesArr = getCSV(ThisWorkbook.path & "\testdata\testtoggleuserverified.csv")
    
    PasteTestRecords testAddressesArr
    
    addRecords
    attemptValidation
    
    AutocorrectedAddressesSheet.Select
    AutocorrectedAddressesSheet.Range("A2:A3").Select
    InterfaceButtons.toggleUserVerifiedAutocorrected
    MacroEntry AutocorrectedAddressesSheet
    
    AutocorrectedAddressesSheet.Range("A2").Select
    InterfaceButtons.toggleUserVerifiedAutocorrected
    MacroEntry InterfaceSheet
    
    CompareSheetCSV Assert, "Addresses", ThisWorkbook.path & "\testdata\testtoggleuserverified_addressesoutput.csv"
    CompareSheetCSV Assert, "Discards", ThisWorkbook.path & "\testdata\testtoggleuserverified_discardsoutput.csv"
    CompareSheetCSV Assert, "Autocorrected", ThisWorkbook.path & "\testdata\testtoggleuserverified_autocorrectedoutput.csv"
    
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestUserMove()
    On Error GoTo TestFail
    
    Dim testAddressesArr() As String
    testAddressesArr = getCSV(ThisWorkbook.path & "\testdata\testusermove.csv")
    PasteTestRecords testAddressesArr
    
    Fakes.MsgBox.Returns vbYes

    InterfaceButtons.confirmAddRecords
    MacroEntry InterfaceSheet
    
    InterfaceButtons.confirmAttemptValidation
    MacroEntry InterfaceSheet
    
    AddressesSheet.Select
    AddressesSheet.Range("A2").Select
    InterfaceButtons.confirmMoveAutocorrect
    MacroEntry InterfaceSheet

    AutocorrectAddressesSheet.Select
    AutocorrectAddressesSheet.Range("A2:A3").Select
    InterfaceButtons.toggleUserVerified
    MacroEntry InterfaceSheet
    
    InterfaceButtons.confirmAttemptValidation
    MacroEntry InterfaceSheet
    
    CompareSheetCSV Assert, "Addresses", ThisWorkbook.path & "\testdata\testusermove_addressesoutput.csv"
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestCountyTotals()
    On Error GoTo TestFail
    
    Dim testAddressesArr() As String
    testAddressesArr = getCSV(ThisWorkbook.path & "\testdata\testcounty.csv")

    PasteTestRecords testAddressesArr

    addRecords
    
    CompareSheetCSV Assert, "Interface", ThisWorkbook.path & "\testdata\testcounty_1added_totalsoutput.csv", getCountyRng
    
    attemptValidation
    
    CompareSheetCSV Assert, "Interface", ThisWorkbook.path & "\testdata\testcounty_2validated_totalsoutput.csv", getCountyRng
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub TestHandcorrected()
    On Error GoTo TestFail
    
    Dim testAddressesArr() As String
'    testAddressesArr = getCSV(ThisWorkbook.path & "\testdata\testhandcorrected.csv")
    testAddressesArr = getCSV(ThisWorkbook.path & "\testdata\testdifficultaddresses.csv")
    PasteTestRecords testAddressesArr
    
    addRecords
    attemptValidation
    
'    CompareSheetCSV Assert, "Addresses", ThisWorkbook.path & "\testdata\testhandcorrected_addressesoutput.csv"
'    CompareSheetCSV Assert, "Needs Autocorrect", ThisWorkbook.path & "\testdata\testhandcorrected_autocorrectoutput.csv"
'    CompareSheetCSV Assert, "Discards", ThisWorkbook.path & "\testdata\testhandcorrected_discardsoutput.csv"
'    CompareSheetCSV Assert, "Autocorrected", ThisWorkbook.path & "\testdata\testhandcorrected_autocorrectedoutput.csv"
    CompareSheetCSV Assert, "Addresses", ThisWorkbook.path & "\testdata\testdifficultaddresses_addressesoutput.csv"
    CompareSheetCSV Assert, "Needs Autocorrect", ThisWorkbook.path & "\testdata\testdifficultaddresses_autocorrectoutput.csv"
    CompareSheetCSV Assert, "Discards", ThisWorkbook.path & "\testdata\testdifficultaddresses_discardsoutput.csv"
    CompareSheetCSV Assert, "Autocorrected", ThisWorkbook.path & "\testdata\testdifficultaddresses_autocorrectedoutput.csv"
    
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestOverwrite()
On Error GoTo TestFail
    
    Dim testAddressesArr() As String
    testAddressesArr = getCSV(ThisWorkbook.path & "\testdata\testoverwrite.csv")
    PasteTestRecords testAddressesArr
    
    addRecords
    attemptValidation
    
    testAddressesArr = getCSV(ThisWorkbook.path & "\testdata\testoverwrite_2.csv")
    PasteTestRecords testAddressesArr
    
    addRecords
    
    CompareSheetCSV Assert, "Addresses", ThisWorkbook.path & "\testdata\testoverwrite_addressesoutput.csv"
    CompareSheetCSV Assert, "Needs Autocorrect", ThisWorkbook.path & "\testdata\testoverwrite_autocorrectoutput.csv"
    CompareSheetCSV Assert, "Discards", ThisWorkbook.path & "\testdata\testoverwrite_discardsoutput.csv"
    CompareSheetCSV Assert, "Autocorrected", ThisWorkbook.path & "\testdata\testoverwrite_autocorrectedoutput.csv"
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestSort()
    On Error GoTo TestFail
    
    Dim wbook As Workbook
    Set wbook = Workbooks.Open(Filename:=ThisWorkbook.path & "\testdata\testsortaddresses.csv")
    
    wbook.Worksheets.[_Default](1).UsedRange.Copy
    AddressesSheet.Range("A1").PasteSpecial xlPasteValues
    
    SheetUtilities.SortSheet "Addresses"
    CompareSheetCSV Assert, "Addresses", ThisWorkbook.path & "\testdata\testsortaddresses_valid_addressesoutput.csv"
    
    SheetUtilities.ClearSheet "Addresses"
    
    wbook.Worksheets.[_Default](1).UsedRange.Copy
    DiscardsSheet.Range("A1").PasteSpecial xlPasteValues
    
    With CreateObject("htmlfile")
        With .parentWindow.clipboardData
            .setData "text", vbNullString
        End With
    End With
    
    wbook.Close
    
    SheetUtilities.SortSheet "Discards"
    CompareSheetCSV Assert, "Discards", ThisWorkbook.path & "\testdata\testsortaddresses_raw_discardsoutput.csv"
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

