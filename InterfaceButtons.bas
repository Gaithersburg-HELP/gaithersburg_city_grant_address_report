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
    
    
    ActiveWorkbook.Worksheets.[_Default]("Interface").Activate
    Application.ScreenUpdating = False
    
    getBlankRow("Interface").Cells.Item(1, 1).Select
    ActiveCell.offset(1, 0).Range("A1").Select
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
                ActiveWorkbook.Worksheets.[_Default]("Addresses").columns(column)
        Else
            Set addressColsToDelete = Union(addressColsToDelete, _
                ActiveWorkbook.Worksheets.[_Default]("Addresses").columns(column))
        End If
        
        Dim service As String
        service = addressServices(column - SheetUtilities.firstServiceColumn)
        
        Dim i As Long
        i = 0
        Do While i <= UBound(autocorrectedServices)
            If service = autocorrectedServices(i) Then
                If autocorrectedColsToDelete Is Nothing Then
                    Set autocorrectedColsToDelete = _
                        ActiveWorkbook.Worksheets.[_Default]("Autocorrected") _
                        .columns(i + SheetUtilities.firstServiceColumn)
                Else
                    Set autocorrectedColsToDelete = Union(autocorrectedColsToDelete, _
                            ActiveWorkbook.Worksheets.[_Default]("Autocorrected") _
                            .columns(i + SheetUtilities.firstServiceColumn))
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
        Set currentRowRng = ActiveWorkbook.Worksheets.[_Default](sourceSheet).Range("A" & row)
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
        ActiveWorkbook.Worksheets.[_Default]("Needs Autocorrect").Cells(row, 2).value = _
            Not ActiveWorkbook.Worksheets.[_Default]("Needs Autocorrect").Cells(row, 2).value
    Next row
End Sub

Public Sub CopyAndOpenCountyTotalsSite()
    ' TODO get current selection row
    
    Dim code As Variant
    code = "var connection = window.indexedDB.open('survey123');"
    code = code & "connection.onsuccess = (e) => {"
    code = code & " var database = e.target.result;"
    code = code & " var tx = database.transaction('data', 'readwrite');"
    code = code & " var objectStore = tx.objectStore(['data']);"
    code = code & " var index = objectStore.index('itemId');"

    code = code & " var request = index.openCursor();"
    code = code & " request.onsuccess = () => {;"
    code = code & "     var cursor = request.result;"
    code = code & "     var fieldJson = cursor.value;"
    
    ' TODO insert totals here
    code = code & "     fieldJson['value']['hh_dup'] = '1';"
    code = code & "     fieldJson['value']['hh_undup'] = '2';"
    code = code & "     fieldJson['value']['individual_dup'] = '3';"
    code = code & "     fieldJson['value']['individual_undup'] = '4';"
    code = code & "     fieldJson['value']['children_ages_0_18'] = '5';"
    code = code & "     fieldJson['value']['adults_ages_18'] = '6';"
    
    ' 20861
    code = code & "     fieldJson['value']['field_7'] = '7';"
    code = code & "     fieldJson['value']['field_8'] = '8';"
    code = code & "     fieldJson['value']['field_9'] = '9';"
    code = code & "     fieldJson['value']['field_14'] = '10';"
    code = code & "     fieldJson['value']['field_11'] = '11';"
    code = code & "     fieldJson['value']['field_15'] = '12';"
    code = code & "     fieldJson['value']['field_16'] = '13';"
    code = code & "     fieldJson['value']['field_17'] = '14';"
    code = code & "     fieldJson['value']['field_18'] = '15';"
    code = code & "     fieldJson['value']['field_19'] = '16';"
    code = code & "     fieldJson['value']['field_20'] = '17';"
    code = code & "     fieldJson['value']['field_21'] = '18';"
    code = code & "     fieldJson['value']['field_22'] = '19';"
    code = code & "     fieldJson['value']['field_23'] = '20';"
    code = code & "     fieldJson['value']['field_24'] = '21';"
    code = code & "     fieldJson['value']['field_25'] = '22';"
    code = code & "     fieldJson['value']['field_26'] = '23';"
    code = code & "     fieldJson['value']['field_27'] = '24';"
    code = code & "     fieldJson['value']['field_28'] = '25';"
    code = code & "     fieldJson['value']['field_29'] = '26';"
    code = code & "     fieldJson['value']['field_30'] = '27';"
    code = code & "     fieldJson['value']['field_31'] = '28';"
    code = code & "     fieldJson['value']['field_32'] = '29';"
    code = code & "     fieldJson['value']['field_37'] = '30';"
    code = code & "     fieldJson['value']['field_35'] = '31';"
    code = code & "     fieldJson['value']['field_36'] = '32';"
    code = code & "     fieldJson['value']['field_34'] = '33';"
    code = code & "     fieldJson['value']['field_38'] = '34';"
    code = code & "     fieldJson['value']['field_39'] = '35';"
    code = code & "     fieldJson['value']['field_40'] = '36';"
    code = code & "     fieldJson['value']['field_41'] = '37';"
    code = code & "     fieldJson['value']['field_42'] = '38';"
    code = code & "     fieldJson['value']['field_43'] = '39';"
    code = code & "     fieldJson['value']['field_44'] = '40';"
    code = code & "     fieldJson['value']['field_45'] = '41';"
    code = code & "     fieldJson['value']['field_46'] = '42';"
    code = code & "     fieldJson['value']['field_47'] = '43';"
    code = code & "     fieldJson['value']['field_48'] = '44';"
    code = code & "     fieldJson['value']['field_49'] = '45';"
    code = code & "     fieldJson['value']['field_50'] = '46';"
    code = code & "     fieldJson['value']['field_51'] = '47';"
    code = code & "     fieldJson['value']['field_52'] = '48';"
    code = code & "     fieldJson['value']['field_53'] = '49';"
    code = code & "     fieldJson['value']['field_54'] = '50';"
    code = code & "     fieldJson['value']['field_55'] = '51';"
    code = code & "     fieldJson['value']['field_56'] = '52';"
    
    code = code & "     request = cursor.update(fieldJson);"
    code = code & "     request.onsuccess = () => {;"
    code = code & "         console.log(request.result);"
    code = code & "         console.log('Fields updated');"
    code = code & "         location.reload();"
    code = code & "     };"
    code = code & " };"
    code = code & "};"

    With CreateObject("htmlfile")
        With .parentWindow.clipboardData
            .setData "text", code
        End With
    End With
    ActiveWorkbook.FollowHyperlink address:="https://survey123.arcgis.com/share/43a57395fe8c4ae5ade7b3bf1e2b8313"
End Sub

' This macro subroutine may be used to double-check
' street addresses by lookup on the Gaithersburg city address search page in browser window.
'@EntryPoint
'@ExcelHotkey L
Public Sub LookupInCity()
Attribute LookupInCity.VB_ProcData.VB_Invoke_Func = "L\n14"
    Dim currentRowFirstCell As Range
    Set currentRowFirstCell = ActiveWorkbook.ActiveSheet.Cells.Item(ActiveCell.row, 1)
    
    Dim record As RecordTuple
    Set record = Records.loadRecordFromSheet(currentRowFirstCell)
    
    Dim AddrLookupURL As String
    AddrLookupURL = "https://maps.gaithersburgmd.gov/AddressSearch/index.html?address="
    AddrLookupURL = AddrLookupURL & record.GburgFormatRawAddress.Item(addressKey.streetAddress)
    AddrLookupURL = Replace(AddrLookupURL, " ", "+")
    
    ActiveWorkbook.FollowHyperlink address:=AddrLookupURL
End Sub

'@EntryPoint
Public Sub OpenAddressValidationWebsite()
    ActiveWorkbook.FollowHyperlink address:="https://developers.google.com/maps/documentation/address-validation/demo"
End Sub

'@EntryPoint
Public Sub OpenUSPSZipcodeWebsite()
    ActiveWorkbook.FollowHyperlink address:="https://tools.usps.com/zip-code-lookup.htm?byaddress"
End Sub

'@EntryPoint
Public Sub OpenGoogleMapsWebsite()
    ActiveWorkbook.FollowHyperlink address:="https://www.google.com/maps"
End Sub
