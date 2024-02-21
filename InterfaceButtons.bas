Attribute VB_Name = "InterfaceButtons"
Option Explicit

'@Folder("City_Grant_Address_Report.src")

Private Function isInvalidSelection() As Boolean
    Dim row As Variant
    For Each row In selection.Rows
        If row < 2 Then
            MsgBox "Invalid Selection"
            isInvalidSelection = True
            Exit Function
        End If
    Next row
    
    isInvalidSelection = False
End Function

'@EntryPoint
Public Sub PasteRecords()
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
    
    autocorrect.attemptValidation
End Sub

'@EntryPoint
Public Sub confirmGenerateFinalReport()
    Dim confirmResponse As VbMsgBoxResult
    confirmResponse = MsgBox("Are you sure you wish to generate the final report?", vbYesNo + vbQuestion, "Confirmation")
    If confirmResponse = vbNo Then
        Exit Sub
    End If
    
    GenerateReport.generateFinalReport
End Sub

'@EntryPoint
Public Sub confirmDeleteAllVisitData()
    Dim confirmResponse As VbMsgBoxResult
    confirmResponse = MsgBox("Are you sure you wish to delete all visit data?", vbYesNo + vbQuestion, "Confirmation")
    If confirmResponse = vbNo Then
        Exit Sub
    End If
    
    SheetUtilities.getAddressVisitDataRng("Addresses").Clear
    SheetUtilities.getAddressVisitDataRng("Needs Autocorrect").Clear
    SheetUtilities.getAddressVisitDataRng("Discards").Clear
    SheetUtilities.getAddressVisitDataRng("Autocorrected").Clear
End Sub

'@EntryPoint
Public Sub confirmDeleteService()
    Dim confirmResponse As VbMsgBoxResult
    confirmResponse = MsgBox("Are you sure you wish to delete the selected service?", vbYesNo + vbQuestion, "Confirmation")
    If confirmResponse = vbNo Then
        Exit Sub
    End If

    'TODO delete selected service
End Sub

'@EntryPoint
Public Sub confirmDiscardAll()
    Dim confirmResponse As VbMsgBoxResult
    confirmResponse = MsgBox("Are you sure you wish to discard all records?", vbYesNo + vbQuestion, "Confirmation")
    If confirmResponse = vbNo Then
        Exit Sub
    End If
    
    Dim autocorrect As Scripting.Dictionary
    Set autocorrect = Records.loadAddresses("Needs Autocorrect")
    
    Dim key As Variant
    For Each key In autocorrect.Keys()
        Records.writeAddress "Discards", autocorrect.Item(key)
    Next key
    
    SheetUtilities.ClearSheet "Needs Autocorrect"
    SheetUtilities.SortSheet "Discards"
End Sub

'@EntryPoint
Public Sub confirmDiscardSelected()
    If isInvalidSelection Then
        Exit Sub
    Else
        Dim confirmResponse As VbMsgBoxResult
        confirmResponse = MsgBox("Are you sure you wish to discard the selected record?", vbYesNo + vbQuestion, "Confirmation")
        If confirmResponse = vbNo Then
            Exit Sub
        End If
    End If
    
    Dim rowsToDelete As Range
    Dim row As Variant
    For Each row In selection.Rows
        Dim record As RecordTuple
        Set record = Records.loadRecordFromSheet(ActiveSheet.Range("A" & row.row))
        Records.writeAddress "Discards", record
        If rowsToDelete Is Nothing Then
            Set rowsToDelete = row
        Else
            Set rowsToDelete = Union(row, rowsToDelete)
        End If
    Next row
    
    rowsToDelete.EntireRow.Delete
    
    ActiveSheet.Cells(1, 1).Select
    SheetUtilities.SortSheet "Discards"
End Sub

'@EntryPoint
Public Sub confirmRestoreSelectedDiscard()
    Dim confirmResponse As VbMsgBoxResult
    confirmResponse = MsgBox("Are you sure you wish to move the selected discard record to Needs Autocorrect?", vbYesNo + vbQuestion, "Confirmation")
    If confirmResponse = vbNo Then
        Exit Sub
    End If
    
    'TODO restore selected
End Sub

'@EntryPoint
Public Sub confirmMoveAutocorrect()
    Dim confirmResponse As VbMsgBoxResult
    confirmResponse = MsgBox("Are you sure you wish to move the selected record to Needs Autocorrect?", vbYesNo + vbQuestion, "Confirmation")
    If confirmResponse = vbNo Then
        Exit Sub
    End If
    ' TODO move to autocorrect
End Sub

'@EntryPoint
Public Sub toggleUserVerified()
    ' TODO toggle User Verified
    If isInvalidSelection Then
    End If
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


