Attribute VB_Name = "InterfaceButtons"
Option Explicit

'@Folder("City_Grant_Address_Report.src")

' Returns Nothing if error occurred
Private Function getUniqueSelection(ByVal returnRows As Boolean, ByVal min As Long) As Collection
    Dim uniques As Collection
    Set uniques = New Collection
    
    Dim dict As Scripting.Dictionary
    Set dict = New Scripting.Dictionary
    
    Dim selections As Range
    ' xlCellTypeVisible in case a filter is applied
    If returnRows Then
        Set selections = selection.SpecialCells(xlCellTypeVisible).rows
    Else
        Set selections = selection.SpecialCells(xlCellTypeVisible).columns
    End If
    
    Dim value As Variant
    For Each value In selections
        If returnRows Then
            If value.row < min Then
                MsgBox "Invalid Selection"
                Set getUniqueSelection = Nothing
                Exit Function
            End If
            dict.Item(value.row) = Empty
        Else
            If value.column < min Then
                MsgBox "Invalid Selection"
                Set getUniqueSelection = Nothing
                Exit Function
            End If
            dict.Item(value.column) = Empty
        End If
    Next value
    
    For Each value In dict.Keys()
        uniques.Add value
    Next value
    
    Set getUniqueSelection = uniques
End Function

'@EntryPoint
Public Sub PasteRecords()
    SheetUtilities.DisableAllFilters
    
    
    InterfaceSheet.Activate
    Application.ScreenUpdating = False
    
    getBlankRow("Interface").Cells.Item(1, 1).Select
    ActiveCell.PasteSpecial (xlPasteValues)
    
    ActiveSheet.Cells(1, 1).Select
    Application.ScreenUpdating = True
End Sub

'@EntryPoint
Public Sub confirmAddRecords()
    Dim confirmResponse As VbMsgBoxResult
    confirmResponse = MsgBox("Are you sure you wish to add records?", vbYesNo + vbQuestion, "Confirmation")
    If confirmResponse = vbNo Then
        Exit Sub
    End If
    
    SheetUtilities.DisableAllFilters
    
    
    Records.addRecords
End Sub

'@EntryPoint
Public Sub confirmAttemptValidation()
    Dim confirmResponse As VbMsgBoxResult
    confirmResponse = MsgBox("Are you sure you wish to attempt validation? You have " & _
                              CStr(getRemainingRequests()) & " remaining.", _
                              vbYesNo + vbQuestion, "Confirmation")
    If confirmResponse = vbNo Then
        Exit Sub
    End If
    
    SheetUtilities.DisableAllFilters
    
    
    autocorrect.attemptValidation
End Sub

'@EntryPoint
Public Sub confirmGenerateFinalReport()
    Dim confirmResponse As VbMsgBoxResult
    confirmResponse = MsgBox("Are you sure you wish to generate the final report?", vbYesNo + vbQuestion, "Confirmation")
    If confirmResponse = vbNo Then
        Exit Sub
    End If
    
    SheetUtilities.DisableAllFilters
    
    
    GenerateReport.generateFinalReport
End Sub

'@EntryPoint
Public Sub confirmDeleteAllVisitData()
    Dim confirmResponse As VbMsgBoxResult
    confirmResponse = MsgBox("Are you sure you wish to delete all visit data?", vbYesNo + vbQuestion, "Confirmation")
    If confirmResponse = vbNo Then
        Exit Sub
    End If
    
    SheetUtilities.DisableAllFilters
    
    SheetUtilities.getTotalsRng.Clear
    ' TODO clear county totals also
    SheetUtilities.getFinalReportRng.Clear
    SheetUtilities.getAddressVisitDataRng("Addresses").Clear
    SheetUtilities.getAddressVisitDataRng("Needs Autocorrect").Clear
    SheetUtilities.getAddressVisitDataRng("Discards").Clear
    SheetUtilities.getAddressVisitDataRng("Autocorrected").Clear
End Sub

'@EntryPoint
Public Sub confirmDeleteService()
    Dim columns As Collection
    Set columns = getUniqueSelection(False, SheetUtilities.firstServiceColumn)
    If columns Is Nothing Then Exit Sub
    
    Dim confirmResponse As VbMsgBoxResult
    confirmResponse = MsgBox("Are you sure you wish to delete the selected service(s)?", vbYesNo + vbQuestion, "Confirmation")
    If confirmResponse = vbNo Then
        Exit Sub
    End If
    
    SheetUtilities.DisableAllFilters
        
    
    Dim addressServices() As String
    addressServices = SheetUtilities.loadServiceNames("Addresses")
    
    Dim autocorrectedServices() As String
    autocorrectedServices = SheetUtilities.loadServiceNames("Autocorrected")
    
    Dim addressColsToDelete As Range
    Dim autocorrectedColsToDelete As Range
    
    Dim column As Variant
    For Each column In columns
        If addressColsToDelete Is Nothing Then
            Set addressColsToDelete = _
                AddressesSheet.columns.Item(column)
        Else
            Set addressColsToDelete = Union(addressColsToDelete, _
                AddressesSheet.columns.Item(column))
        End If
        
        Dim service As String
        service = addressServices(column - SheetUtilities.firstServiceColumn)
        
        Dim i As Long
        i = 0
        Do While i <= UBound(autocorrectedServices)
            If service = autocorrectedServices(i) Then
                If autocorrectedColsToDelete Is Nothing Then
                    Set autocorrectedColsToDelete = _
                        AutocorrectedAddressesSheet _
                        .columns.Item(i + SheetUtilities.firstServiceColumn)
                Else
                    Set autocorrectedColsToDelete = Union(autocorrectedColsToDelete, _
                            AutocorrectedAddressesSheet _
                            .columns.Item(i + SheetUtilities.firstServiceColumn))
                End If
                Exit Do
            End If
            i = i + 1
        Loop
    Next column
    
    addressColsToDelete.EntireColumn.Delete
    
    If Not autocorrectedColsToDelete Is Nothing Then
        autocorrectedColsToDelete.EntireColumn.Delete
    End If
    
    SheetUtilities.getFinalReportRng.Clear
    Records.computeTotals
    Records.computeCountyTotals
End Sub

'@EntryPoint
Public Sub confirmDiscardAll()
    Dim confirmResponse As VbMsgBoxResult
    confirmResponse = MsgBox("Are you sure you wish to discard all records?", vbYesNo + vbQuestion, "Confirmation")
    If confirmResponse = vbNo Then
        Exit Sub
    End If
    
    SheetUtilities.DisableAllFilters
    
    
    Dim autocorrect As Scripting.Dictionary
    Set autocorrect = Records.loadAddresses("Needs Autocorrect")
    
    Dim key As Variant
    For Each key In autocorrect.Keys()
        Records.writeAddress "Discards", autocorrect.Item(key)
    Next key
    
    SheetUtilities.ClearSheet "Needs Autocorrect"
    SheetUtilities.SortSheet "Discards"
End Sub

Private Sub moveSelectedRows(ByVal sourceSheet As String, ByVal destSheet As String, _
                             ByVal removeFromAutocorrected As Boolean)
    Dim rows As Collection
    Set rows = getUniqueSelection(True, 2)
    If rows Is Nothing Then
        Exit Sub
    End If
    
    SheetUtilities.DisableAllFilters
    
    
    Dim confirmResponse As VbMsgBoxResult
    confirmResponse = MsgBox("Are you sure you wish to move the selected record(s) from " & _
                             sourceSheet & " to " & destSheet & "?", vbYesNo + vbQuestion, "Confirmation")
    If confirmResponse = vbNo Then
        Exit Sub
    End If
    
    Dim movedRecords As Collection
    Set movedRecords = New Collection
    
    Dim rowsToDelete As Range
    Dim row As Variant
    For Each row In rows
        Dim currentRowRng As Range
        Set currentRowRng = ThisWorkbook.Worksheets.[_Default](sourceSheet).Range("A" & row)
        Dim record As RecordTuple
        Set record = Records.loadRecordFromSheet(currentRowRng)
        
        Records.writeAddress destSheet, record
        movedRecords.Add record
        
        If rowsToDelete Is Nothing Then
            Set rowsToDelete = currentRowRng
        Else
            Set rowsToDelete = Union(currentRowRng, rowsToDelete)
        End If
    Next row
    
    rowsToDelete.EntireRow.Delete
    SheetUtilities.ClearEmptyServices sourceSheet
    
    ActiveSheet.Cells(1, 1).Select
    SheetUtilities.SortSheet destSheet
    
    If (Not removeFromAutocorrected) Then Exit Sub
    
    Dim autocorrected As Scripting.Dictionary
    Set autocorrected = Records.loadAddresses("Autocorrected")
    
    Dim changedAutocorrected As Boolean
    changedAutocorrected = False
    
    Dim movedRecord As Variant
    For Each movedRecord In movedRecords
        If autocorrected.Exists(movedRecord.key) Then
            changedAutocorrected = True
            autocorrected.Remove movedRecord.key
        End If
    Next movedRecord
    
    If changedAutocorrected Then Records.writeAddresses "Autocorrected", autocorrected
End Sub

'@EntryPoint
Public Sub confirmDiscardSelected()
    moveSelectedRows "Needs Autocorrect", "Discards", False
End Sub

'@EntryPoint
Public Sub confirmRestoreSelectedDiscard()
    moveSelectedRows "Discards", "Needs Autocorrect", True
End Sub

'@EntryPoint
Public Sub confirmMoveAutocorrect()
    moveSelectedRows "Addresses", "Needs Autocorrect", True
    SheetUtilities.getFinalReportRng.Clear
    Records.computeTotals
End Sub

'@EntryPoint
Public Sub toggleUserVerified()
    Dim rows As Collection
    Set rows = getUniqueSelection(True, 2)
    
    If rows Is Nothing Then Exit Sub
    
    Dim row As Variant
    For Each row In rows
        AutocorrectAddressesSheet.Cells.Item(row, 2).value = _
            Not AutocorrectAddressesSheet.Cells.Item(row, 2).value
    Next row
End Sub

'@EntryPoint
Public Sub toggleUserVerifiedAutocorrected()
    Dim rows As Collection
    Set rows = getUniqueSelection(True, 2)
    
    If rows Is Nothing Then Exit Sub
    
    Dim addresses As Scripting.Dictionary
    Set addresses = Records.loadAddresses("Addresses")
    
    Dim autocorrected As Scripting.Dictionary
    Set autocorrected = Records.loadAddresses("Autocorrected")
    
    Dim discards As Scripting.Dictionary
    Set discards = Records.loadAddresses("Discards")
    
    Dim addressesModified As Boolean
    Dim discardsModified As Boolean
    
    Dim row As Variant
    For Each row In rows
        Dim currentRowRng As Range
        Set currentRowRng = AutocorrectedAddressesSheet.Range("A" & row)
        Dim record As RecordTuple
        Set record = Records.loadRecordFromSheet(currentRowRng)
        
        autocorrected.Item(record.key).UserVerified = Not autocorrected.Item(record.key).UserVerified
        
        If addresses.Exists(record.key) Then
            addresses.Item(record.key).UserVerified = Not addresses.Item(record.key).UserVerified
            addressesModified = True
        ElseIf discards.Exists(record.key) Then
            discards.Item(record.key).UserVerified = Not discards.Item(record.key).UserVerified
            discardsModified = True
        End If
    Next row
    
    SheetUtilities.ClearSheet "Autocorrected"
    Records.writeAddresses "Autocorrected", autocorrected
    
    If addressesModified Then
        SheetUtilities.ClearSheet "Addresses"
        Records.writeAddresses "Addresses", addresses
    End If
    If discardsModified Then
        SheetUtilities.ClearSheet "Discards"
        Records.writeAddresses "Discards", discards
    End If
End Sub

'@EntryPoint
Public Sub CopyAndOpenCountyTotalsSite()
    Dim values As Range
    Set values = ActiveSheet.rows(selection.row)
    
    Dim code As Variant
    code = vbNullString
    code = code & "var connection = window.indexedDB.open('survey123');"
    code = code & "connection.onsuccess = (e) => {"
    code = code & "    const database = e.target.result;"
    code = code & "    const tx = database.transaction('data', 'readwrite');"
    code = code & "    const objectStore = tx.objectStore(['data']);"
    code = code & "    const index = objectStore.index('itemId');"
    code = code & "    let request = index.openCursor();"
    code = code & "    request.onsuccess = () => {"
    code = code & "        const cursor = request.result;"
    code = code & "        const fieldJson = cursor.value;"
    code = code & "        fieldJson['value']['month_report'] = '" & values.Cells.Item(1, 1).value & "';"
    code = code & "        fieldJson['value']['hh_dup'] ='" & values.Cells.Item(1, 2).value & "';"
    code = code & "        fieldJson['value']['hh_undup'] ='" & values.Cells.Item(1, 3).value & "';"
    code = code & "        fieldJson['value']['individual_dup'] ='" & values.Cells.Item(1, 4).value & "';"
    code = code & "        fieldJson['value']['individual_undup'] ='" & values.Cells.Item(1, 5).value & "';"
    code = code & "        fieldJson['value']['children_ages_0_18'] ='" & values.Cells.Item(1, 6).value & "';"
    code = code & "        fieldJson['value']['adults_ages_18'] ='" & values.Cells.Item(1, 7).value & "';"
    code = code & "        fieldJson['value']['fa_pre_pack_boxbags'] ='" & values.Cells.Item(1, 8).value & "';"
    code = code & "        fieldJson['value']['field_7'] ='" & values.Cells.Item(1, 9).value & "';"
    code = code & "        fieldJson['value']['field_8'] ='" & values.Cells.Item(1, 10).value & "';"
    code = code & "        fieldJson['value']['field_9'] ='" & values.Cells.Item(1, 11).value & "';"
    code = code & "        fieldJson['value']['field_14'] ='" & values.Cells.Item(1, 12).value & "';"
    code = code & "        fieldJson['value']['field_11'] ='" & values.Cells.Item(1, 13).value & "';"
    code = code & "        fieldJson['value']['field_15'] ='" & values.Cells.Item(1, 14).value & "';"
    code = code & "        fieldJson['value']['field_16'] ='" & values.Cells.Item(1, 15).value & "';"
    code = code & "        fieldJson['value']['field_17'] ='" & values.Cells.Item(1, 16).value & "';"
    code = code & "        fieldJson['value']['field_18'] ='" & values.Cells.Item(1, 17).value & "';"
    code = code & "        fieldJson['value']['field_19'] ='" & values.Cells.Item(1, 18).value & "';"
    code = code & "        fieldJson['value']['field_20'] ='" & values.Cells.Item(1, 19).value & "';"
    code = code & "        fieldJson['value']['field_21'] ='" & values.Cells.Item(1, 20).value & "';"
    code = code & "        fieldJson['value']['field_22'] ='" & values.Cells.Item(1, 21).value & "';"
    code = code & "        fieldJson['value']['field_106'] ='" & values.Cells.Item(1, 22).value & "';"
    code = code & "        fieldJson['value']['field_23'] ='" & values.Cells.Item(1, 23).value & "';"
    code = code & "        fieldJson['value']['field_24'] ='" & values.Cells.Item(1, 24).value & "';"
    code = code & "        fieldJson['value']['field_25'] ='" & values.Cells.Item(1, 25).value & "';"
    code = code & "        fieldJson['value']['field_26'] ='" & values.Cells.Item(1, 26).value & "';"
    code = code & "        fieldJson['value']['field_27'] ='" & values.Cells.Item(1, 27).value & "';"
    code = code & "        fieldJson['value']['field_28'] ='" & values.Cells.Item(1, 28).value & "';"
    code = code & "        fieldJson['value']['field_29'] ='" & values.Cells.Item(1, 29).value & "';"
    code = code & "        fieldJson['value']['field_30'] ='" & values.Cells.Item(1, 30).value & "';"
    code = code & "        fieldJson['value']['field_31'] ='" & values.Cells.Item(1, 31).value & "';"
    code = code & "        fieldJson['value']['field_32'] ='" & values.Cells.Item(1, 32).value & "';"
    code = code & "        fieldJson['value']['field_37'] ='" & values.Cells.Item(1, 33).value & "';"
    code = code & "        fieldJson['value']['field_35'] ='" & values.Cells.Item(1, 34).value & "';"
    code = code & "        fieldJson['value']['field_36'] ='" & values.Cells.Item(1, 35).value & "';"
    code = code & "        fieldJson['value']['field_34'] ='" & values.Cells.Item(1, 36).value & "';"
    code = code & "        fieldJson['value']['field_38'] ='" & values.Cells.Item(1, 37).value & "';"
    code = code & "        fieldJson['value']['field_39'] ='" & values.Cells.Item(1, 38).value & "';"
    code = code & "        fieldJson['value']['field_40'] ='" & values.Cells.Item(1, 39).value & "';"
    code = code & "        fieldJson['value']['field_41'] ='" & values.Cells.Item(1, 40).value & "';"
    code = code & "        fieldJson['value']['field_42'] ='" & values.Cells.Item(1, 41).value & "';"
    code = code & "        fieldJson['value']['field_43'] ='" & values.Cells.Item(1, 42).value & "';"
    code = code & "        fieldJson['value']['field_44'] ='" & values.Cells.Item(1, 43).value & "';"
    code = code & "        fieldJson['value']['field_45'] ='" & values.Cells.Item(1, 44).value & "';"
    code = code & "        fieldJson['value']['field_46'] ='" & values.Cells.Item(1, 45).value & "';"
    code = code & "        fieldJson['value']['field_47'] ='" & values.Cells.Item(1, 46).value & "';"
    code = code & "        fieldJson['value']['field_48'] ='" & values.Cells.Item(1, 47).value & "';"
    code = code & "        fieldJson['value']['field_49'] ='" & values.Cells.Item(1, 48).value & "';"
    code = code & "        fieldJson['value']['field_50'] ='" & values.Cells.Item(1, 49).value & "';"
    code = code & "        fieldJson['value']['field_51'] ='" & values.Cells.Item(1, 50).value & "';"
    code = code & "        fieldJson['value']['field_52'] ='" & values.Cells.Item(1, 51).value & "';"
    code = code & "        fieldJson['value']['field_53'] ='" & values.Cells.Item(1, 52).value & "';"
    code = code & "        fieldJson['value']['field_54'] ='" & values.Cells.Item(1, 53).value & "';"
    code = code & "        fieldJson['value']['field_55'] ='" & values.Cells.Item(1, 54).value & "';"
    code = code & "        fieldJson['value']['field_56'] ='" & values.Cells.Item(1, 55).value & "';"
    code = code & "        fieldJson['value']['field_107'] ='" & values.Cells.Item(1, 56).value & "';"
    code = code & "        fieldJson['value']['field_108'] ='" & values.Cells.Item(1, 57).value & "';"
    code = code & "        fieldJson['value']['field_109'] ='" & values.Cells.Item(1, 58).value & "';"
    code = code & "        fieldJson['value']['field_110'] ='" & values.Cells.Item(1, 59).value & "';"
    code = code & "        fieldJson['value']['field_111'] ='" & values.Cells.Item(1, 60).value & "';"
    code = code & "        fieldJson['value']['field_112'] ='" & values.Cells.Item(1, 61).value & "';"
    code = code & "        fieldJson['value']['field_113'] ='" & values.Cells.Item(1, 62).value & "';"
    code = code & "        fieldJson['value']['field_114'] ='" & values.Cells.Item(1, 63).value & "';"
    code = code & "        fieldJson['value']['field_115'] ='" & values.Cells.Item(1, 64).value & "';"
    code = code & "        fieldJson['value']['field_116'] ='" & values.Cells.Item(1, 65).value & "';"
    code = code & "        fieldJson['value']['field_117'] ='" & values.Cells.Item(1, 66).value & "';"
    code = code & "        fieldJson['value']['field_118'] ='" & values.Cells.Item(1, 67).value & "';"
    code = code & "        fieldJson['value']['field_119'] ='" & values.Cells.Item(1, 68).value & "';"
    code = code & "        fieldJson['value']['field_120'] ='" & values.Cells.Item(1, 69).value & "';"
    code = code & "        fieldJson['value']['field_121'] ='" & values.Cells.Item(1, 70).value & "';"
    code = code & "        fieldJson['value']['field_122'] ='" & values.Cells.Item(1, 71).value & "';"
    code = code & "        fieldJson['value']['field_123'] ='" & values.Cells.Item(1, 72).value & "';"
    code = code & "        fieldJson['value']['field_124'] ='" & values.Cells.Item(1, 73).value & "';"
    code = code & "        fieldJson['value']['field_125'] ='" & values.Cells.Item(1, 74).value & "';"
    code = code & "        fieldJson['value']['field_126'] ='" & values.Cells.Item(1, 75).value & "';"
    code = code & "        fieldJson['value']['field_127'] ='" & values.Cells.Item(1, 76).value & "';"
    code = code & "        fieldJson['value']['field_128'] ='" & values.Cells.Item(1, 77).value & "';"
    code = code & "        fieldJson['value']['field_129'] ='" & values.Cells.Item(1, 78).value & "';"
    code = code & "        fieldJson['value']['field_130'] ='" & values.Cells.Item(1, 79).value & "';"
    code = code & "        fieldJson['value']['field_131'] ='" & values.Cells.Item(1, 80).value & "';"
    code = code & "        fieldJson['value']['field_132'] ='" & values.Cells.Item(1, 81).value & "';"
    code = code & "        fieldJson['value']['field_133'] ='" & values.Cells.Item(1, 82).value & "';"
    code = code & "        fieldJson['value']['field_134'] ='" & values.Cells.Item(1, 83).value & "';"
    code = code & "        fieldJson['value']['field_135'] ='" & values.Cells.Item(1, 84).value & "';"
    code = code & "        fieldJson['value']['field_136'] ='" & values.Cells.Item(1, 85).value & "';"
    code = code & "        fieldJson['value']['field_137'] ='" & values.Cells.Item(1, 86).value & "';"
    code = code & "        fieldJson['value']['field_138'] ='" & values.Cells.Item(1, 87).value & "';"
    code = code & "        fieldJson['value']['field_139'] ='" & values.Cells.Item(1, 88).value & "';"
    code = code & "        fieldJson['value']['field_140'] ='" & values.Cells.Item(1, 89).value & "';"
    code = code & "        fieldJson['value']['field_141'] ='" & values.Cells.Item(1, 90).value & "';"
    code = code & "        fieldJson['value']['field_142'] ='" & values.Cells.Item(1, 91).value & "';"
    code = code & "        fieldJson['value']['field_143'] ='" & values.Cells.Item(1, 92).value & "';"
    code = code & "        fieldJson['value']['field_144'] ='" & values.Cells.Item(1, 93).value & "';"
    code = code & "        fieldJson['value']['field_145'] ='" & values.Cells.Item(1, 94).value & "';"
    code = code & "        fieldJson['value']['field_146'] ='" & values.Cells.Item(1, 95).value & "';"
    code = code & "        fieldJson['value']['field_147'] ='" & values.Cells.Item(1, 96).value & "';"
    code = code & "        fieldJson['value']['field_148'] ='" & values.Cells.Item(1, 97).value & "';"
    code = code & "        request = cursor.update(fieldJson);"
    code = code & "        request.onsuccess = () => {"
    code = code & "            location.reload();"
    code = code & "        }"
    code = code & "    }"
    code = code & "}"

    With CreateObject("htmlfile")
        With .parentWindow.clipboardData
            .setData "text", code
        End With
    End With
    ThisWorkbook.FollowHyperlink address:="https://survey123.arcgis.com/share/43a57395fe8c4ae5ade7b3bf1e2b8313"
End Sub

' This macro subroutine may be used to double-check
' street addresses by lookup on the Gaithersburg city address search page in browser window.
'@EntryPoint
'@ExcelHotkey L
Public Sub LookupInCity()
Attribute LookupInCity.VB_ProcData.VB_Invoke_Func = "L\n14"
    Dim currentRowFirstCell As Range
    Set currentRowFirstCell = ThisWorkbook.ActiveSheet.Cells.Item(ActiveCell.row, 1)
    
    Dim record As RecordTuple
    Set record = Records.loadRecordFromSheet(currentRowFirstCell)
    
    Dim AddrLookupURL As String
    AddrLookupURL = "https://maps.gaithersburgmd.gov/AddressSearch/index.html?address="
    AddrLookupURL = AddrLookupURL & record.GburgFormatRawAddress.Item(addressKey.streetAddress)
    AddrLookupURL = Replace(AddrLookupURL, " ", "+")
    
    ThisWorkbook.FollowHyperlink address:=AddrLookupURL
End Sub

'@EntryPoint
Public Sub OpenAddressValidationWebsite()
    ThisWorkbook.FollowHyperlink address:="https://developers.google.com/maps/documentation/address-validation/demo"
End Sub

'@EntryPoint
Public Sub OpenUSPSZipcodeWebsite()
    ThisWorkbook.FollowHyperlink address:="https://tools.usps.com/zip-code-lookup.htm?byaddress"
End Sub

'@EntryPoint
Public Sub OpenGoogleMapsWebsite()
    ThisWorkbook.FollowHyperlink address:="https://www.google.com/maps"
End Sub
