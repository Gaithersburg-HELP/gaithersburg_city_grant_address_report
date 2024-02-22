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

Private Function moveSelectedRows(ByVal sourceSheet As String, ByVal destSheet As String) As Collection
    Dim rows As Collection
    Set rows = getUniqueSelection(True, 2)
    If rows Is Nothing Then
        Set moveSelectedRows = Nothing
        Exit Function
    End If
    
    SheetUtilities.DisableAllFilters
    
    
    Dim confirmResponse As VbMsgBoxResult
    confirmResponse = MsgBox("Are you sure you wish to move the selected record(s) from " & _
                             sourceSheet & " to " & destSheet & "?", vbYesNo + vbQuestion, "Confirmation")
    If confirmResponse = vbNo Then
        Set moveSelectedRows = Nothing
        Exit Function
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
    
    Set moveSelectedRows = movedRecords
End Function

'@EntryPoint
Public Sub confirmDiscardSelected()
    '@Ignore FunctionReturnValueDiscarded
    moveSelectedRows "Needs Autocorrect", "Discards"
End Sub

'@EntryPoint
Public Sub confirmRestoreSelectedDiscard()
    '@Ignore FunctionReturnValueDiscarded
    moveSelectedRows "Discards", "Needs Autocorrect"
End Sub

'@EntryPoint
Public Sub confirmMoveAutocorrect()
    Dim movedRecords As Collection
    Set movedRecords = moveSelectedRows("Addresses", "Needs Autocorrect")
    
    If movedRecords Is Nothing Then Exit Sub
    
    
    Dim autocorrected As Scripting.Dictionary
    Set autocorrected = Records.loadAddresses("Autocorrected")
    
    Dim changedAutocorrected As Boolean
    changedAutocorrected = False
    
    Dim record As Variant
    For Each record In movedRecords
        If autocorrected.Exists(record.key) Then
            changedAutocorrected = True
            autocorrected.Remove record.key
        End If
    Next record
    
    If changedAutocorrected Then Records.writeAddresses "Autocorrected", autocorrected
End Sub

'@EntryPoint
Public Sub toggleUserVerified()
    Dim rows As Collection
    Set rows = getUniqueSelection(True, 2)
    
    If rows Is Nothing Then Exit Sub
    
    Dim row As Variant
    For Each row In rows
        ActiveSheet.Cells(row, 2).value = Not ActiveSheet.Cells(row, 2)
    Next row
End Sub

' This macro subroutine may be used to double-check
' street addresses by lookup on the Gaithersburg city address search page in browser window.
'@EntryPoint
Public Sub LookupInCity()
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


