Attribute VB_Name = "Records"
Option Explicit

'@Folder "City_Grant_Address_Report.src"
'@EntryPoint
Public Sub confirmAddRecords()
    Dim confirmResponse As VbMsgBoxResult
    confirmResponse = MsgBox("Are you sure you wish to add records?", vbYesNo + vbQuestion, "Confirmation")
    If confirmResponse = vbNo Then
        Exit Sub
    End If
    
    addRecords
End Sub

Public Function getQuarterNum(ByVal quarter As String) As Long
    Select Case quarter
        Case "Q1"
            getQuarterNum = 1
        Case "Q2"
            getQuarterNum = 2
        Case "Q3"
            getQuarterNum = 3
        Case "Q4"
            getQuarterNum = 4
    End Select
End Function

Private Function loadRecordFromRaw(ByVal recordRowFirstCell As Range) As RecordTuple
    Dim record As RecordTuple
    Set record = New RecordTuple
    
    record.AddVisit recordRowFirstCell.Value, recordRowFirstCell.Offset(0, 1).Value
    record.UserVerified = False

    record.guestID = recordRowFirstCell.Offset(0, 2).Value
    record.FirstName = recordRowFirstCell.Offset(0, 3).Value
    record.LastName = recordRowFirstCell.Offset(0, 4).Value
    record.RawAddress = recordRowFirstCell.Offset(0, 5).Value
    record.RawUnitWithNum = recordRowFirstCell.Offset(0, 6).Value
    record.RawCity = recordRowFirstCell.Offset(0, 7).Value
    record.RawState = recordRowFirstCell.Offset(0, 8).Value
    record.RawZip = recordRowFirstCell.Offset(0, 9).Value
    record.householdTotal = recordRowFirstCell.Offset(0, 10).Value
    
    Dim rx As Double
    rx = recordRowFirstCell.Offset(0, 11).Value
    If rx <> 0 Then record.addRx recordRowFirstCell.Value, rx
    
    Set loadRecordFromRaw = record
End Function

Public Function loadRecordFromSheet(ByVal recordRowFirstCell As Range) As RecordTuple
    Dim record As RecordTuple
    Set record = New RecordTuple
    
    Dim services() As String
    services = loadServiceNames(recordRowFirstCell.Worksheet.name)
    
    record.SetInCity recordRowFirstCell.Offset(0, 0).Value
    record.UserVerified = CBool(recordRowFirstCell.Offset(0, 1).Value)
    record.ValidAddress = recordRowFirstCell.Offset(0, 2).Value
    record.validUnitWithNum = recordRowFirstCell.Offset(0, 3).Value
    record.ValidZipcode = recordRowFirstCell.Offset(0, 4).Value
    record.RawAddress = recordRowFirstCell.Offset(0, 5).Value
    record.RawUnitWithNum = recordRowFirstCell.Offset(0, 6).Value
    record.RawCity = recordRowFirstCell.Offset(0, 7).Value
    record.RawState = recordRowFirstCell.Offset(0, 8).Value
    record.RawZip = recordRowFirstCell.Offset(0, 9).Value
    record.guestID = recordRowFirstCell.Offset(0, 10).Value
    record.FirstName = recordRowFirstCell.Offset(0, 11).Value
    record.LastName = recordRowFirstCell.Offset(0, 12).Value
    record.householdTotal = recordRowFirstCell.Offset(0, 13).Value
    Set record.rxTotal = JsonConverter.ParseJson(recordRowFirstCell.Offset(0, 14).Value)
    
    Dim visitData As Scripting.Dictionary
    Set visitData = New Scripting.Dictionary
    
    Dim j As Long
    j = 1
    Do While j <= UBound(services) + 1
        Dim visitJson As String
        visitJson = recordRowFirstCell.Offset(0, 14 + j).Value
        If visitJson <> vbNullString Then
            visitData.Add services(j - 1), JsonConverter.ParseJson(visitJson)
        End If
        j = j + 1
    Loop
    
    Set record.visitData = visitData
    
    Set loadRecordFromSheet = record
End Function

Public Function loadAddresses(ByVal sheetName As String) As Scripting.Dictionary
    Dim addresses As Scripting.Dictionary
    Set addresses = New Scripting.Dictionary
    
    Dim sheet As Worksheet
    Set sheet = ActiveWorkbook.Worksheets.[_Default](sheetName)
       
    If sheet.Range("A2").Value = vbNullString Then
        Set loadAddresses = addresses
        Exit Function
    End If
    
    Dim i As Long
    i = 2
    Do While i < getBlankRow(sheetName).row
        Dim recordRowFirstCell As Range
        Set recordRowFirstCell = sheet.Rows.Item(i).Cells.Item(1, 1)
        
        Dim record As RecordTuple
        Set record = loadRecordFromSheet(recordRowFirstCell)
        
        addresses.Add record.key, record
        i = i + 1
    Loop

    Set loadAddresses = addresses
End Function

Public Sub writeAddress(ByVal sheetName As String, ByVal record As RecordTuple)
    Dim sheet As Worksheet
    Set sheet = ActiveWorkbook.Worksheets.[_Default](sheetName)
    
    ' Saves column numbers per existing service
    Dim serviceCols As Scripting.Dictionary
    Set serviceCols = New Scripting.Dictionary
    
    If sheet.Range("A2").Value <> vbNullString Then
        Dim services() As String
        services = loadServiceNames(sheetName)
        Dim i As Long
        i = 16
        Dim service As Variant
        For Each service In services
            serviceCols.Add service, i
            i = i + 1
        Next
    End If

    Dim recordRow As Range
    Set recordRow = getBlankRow(sheetName)
    
    recordRow.Cells.Item(1, 1).Value = record.InCityStr
    recordRow.Cells.Item(1, 2).Value = record.UserVerified
    recordRow.Cells.Item(1, 3).Value = record.ValidAddress
    recordRow.Cells.Item(1, 4).Value = record.validUnitWithNum
    recordRow.Cells.Item(1, 5).Value = record.ValidZipcode
    recordRow.Cells.Item(1, 6).Value = record.RawAddress
    recordRow.Cells.Item(1, 7).Value = record.RawUnitWithNum
    recordRow.Cells.Item(1, 8).Value = record.RawCity
    recordRow.Cells.Item(1, 9).Value = record.RawState
    recordRow.Cells.Item(1, 10).Value = record.RawZip
    recordRow.Cells.Item(1, 11).Value = record.guestID
    recordRow.Cells.Item(1, 12).Value = record.FirstName
    recordRow.Cells.Item(1, 13).Value = record.LastName
    recordRow.Cells.Item(1, 14).Value = record.householdTotal
    
    recordRow.Cells.Item(1, 15).Value = JsonConverter.ConvertToJson(record.rxTotal)
    
    Dim serviceToAdd As Variant
    For Each serviceToAdd In record.visitData.Keys
        Dim visitDataToAdd As String
        visitDataToAdd = JsonConverter.ConvertToJson(record.visitData.Item(serviceToAdd))
        
        If Not serviceCols.Exists(serviceToAdd) Then
            Dim newServiceCol As Long
            newServiceCol = 15 + 1 + UBound(serviceCols.Keys) + 1
            serviceCols.Add serviceToAdd, newServiceCol
            ActiveWorkbook.Worksheets.[_Default](sheetName).Cells(1, newServiceCol).Value = serviceToAdd
        End If
        
        recordRow.Cells.Item(1, serviceCols.Item(serviceToAdd)).Value = visitDataToAdd
    Next serviceToAdd
End Sub

Public Sub writeAddresses(ByVal sheetName As String, ByVal addresses As Scripting.Dictionary)
    ClearSheet sheetName
    Dim key As Variant
    For Each key In addresses.Keys
        writeAddress sheetName, addresses.Item(key)
    Next key
End Sub

Public Sub writeAddressesComputeTotals(ByVal addresses As Scripting.Dictionary, _
                                       ByVal needsAutocorrect As Scripting.Dictionary, _
                                       ByVal discards As Scripting.Dictionary, _
                                       ByVal autocorrected As Scripting.Dictionary)
    SheetUtilities.ClearAll
    
    ' All initialized to 0
    Dim uniqueGuestIDTotal(1 To 4) As Long
    Dim uniqueGuestIDHouseholdTotal(1 To 4) As Long
    Dim guestIDTotal(1 To 4) As Long
    Dim householdTotal(1 To 4) As Long
    Dim rxTotal(1 To 4) As Double
    
    Dim key As Variant
    For Each key In addresses.Keys
        Dim record As RecordTuple
        Set record = addresses.Item(key)
        writeAddress "Addresses", addresses.Item(key)
        
        If record.InCity = InCityCode.ValidInCity Then
            Dim rxCount(1 To 4) As Double
            Dim quarter As Variant
            For Each quarter In record.rxTotal.Keys
                Dim visit As Variant
                For Each visit In record.rxTotal.Item(quarter).Keys
                    rxCount(getQuarterNum(quarter)) = rxCount(getQuarterNum(quarter)) + _
                                                      record.rxTotal.Item(quarter).Item(visit)
                Next visit
            Next quarter
            
            Dim visitCount(1 To 4) As Long
            
            Dim service As Variant
            For Each service In record.visitData.Keys
                For Each quarter In record.visitData.Item(service).Keys
                    visitCount(getQuarterNum(quarter)) = _
                        visitCount(getQuarterNum(quarter)) + _
                        record.visitData.Item(service).Item(quarter).Count
                Next quarter
            Next service
            
            Dim i As Long
            For i = 1 To 4
                If visitCount(i) > 0 Then
                    uniqueGuestIDTotal(i) = uniqueGuestIDTotal(i) + 1
                    uniqueGuestIDHouseholdTotal(i) = uniqueGuestIDHouseholdTotal(i) + _
                                                     record.householdTotal
                End If
                guestIDTotal(i) = guestIDTotal(i) + visitCount(i)
                householdTotal(i) = householdTotal(i) + (visitCount(i) * record.householdTotal)
                rxTotal(i) = rxTotal(i) + rxCount(i)
                
                ' arrays are not reset on loop iteration!
                rxCount(i) = 0
                visitCount(i) = 0
            Next i
        End If
    Next key
    
    Dim totalsRng As Range
    Set totalsRng = SheetUtilities.getTotalsRng
    
    For i = 1 To 4
        totalsRng.Cells.Item(1, i) = uniqueGuestIDTotal(i)
        totalsRng.Cells.Item(2, i) = uniqueGuestIDHouseholdTotal(i)
        totalsRng.Cells.Item(3, i) = guestIDTotal(i)
        totalsRng.Cells.Item(4, i) = householdTotal(i)
        totalsRng.Cells.Item(5, i) = rxTotal(i)
    Next i
    
    writeAddresses "Needs Autocorrect", needsAutocorrect
    writeAddresses "Discards", discards
    writeAddresses "Autocorrected", autocorrected
    
    SortAll
End Sub

Public Sub addRecords()
    ' TODO import MicroTimer from Module 1
    ' Save application status bar to restore it later
    Dim appStatus As Variant
    If Application.StatusBar = False Then appStatus = False Else appStatus = Application.StatusBar
    
    Application.StatusBar = "Loading addresses"
        
    Dim addresses As Scripting.Dictionary
    Set addresses = loadAddresses("Addresses")
    
    Dim needsAutocorrect As Scripting.Dictionary
    Set needsAutocorrect = loadAddresses("Needs Autocorrect")
    
    Dim discards As Scripting.Dictionary
    Set discards = loadAddresses("Discards")
    
    Dim autocorrected As Scripting.Dictionary
    Set autocorrected = loadAddresses("Autocorrected")
       
    Dim recordsToValidate As Scripting.Dictionary
    Set recordsToValidate = New Scripting.Dictionary
    
    Dim i As Long
    i = 9
    Do While i < getBlankRow("Interface").row
        Dim recordToAdd As RecordTuple
        Set recordToAdd = loadRecordFromRaw(ActiveWorkbook.Sheets.[_Default]("Interface").Range("A" & i))
        
        Dim existingRecord As RecordTuple
        
        If addresses.Exists(recordToAdd.key) Then
            Set existingRecord = addresses.Item(recordToAdd.key)
            existingRecord.MergeRecord recordToAdd
            If autocorrected.Exists(recordToAdd.key) Then
                Set existingRecord = autocorrected.Item(recordToAdd.key)
                existingRecord.MergeRecord recordToAdd
            End If
        ElseIf needsAutocorrect.Exists(recordToAdd.key) Then
            Set existingRecord = needsAutocorrect.Item(recordToAdd.key)
            existingRecord.MergeRecord recordToAdd
        ElseIf discards.Exists(recordToAdd.key) Then
            ' BUG if previously discarded user ID but address was updated, this will be discarded anyway
            Set existingRecord = discards.Item(recordToAdd.key)
            existingRecord.MergeRecord recordToAdd
        ElseIf recordsToValidate.Exists(recordToAdd.key) Then
            Set existingRecord = recordsToValidate.Item(recordToAdd.key)
            existingRecord.MergeRecord recordToAdd
        Else
            If recordToAdd.isCorrectableAddress() Then
                recordsToValidate.Add recordToAdd.key, recordToAdd
            Else
                recordToAdd.SetInCity InCityCode.NotCorrectable
                discards.Add recordToAdd.key, recordToAdd
            End If
        End If
        
        Application.StatusBar = "Adding record " & (i - 8) & " of " & (getBlankRow("Interface").row - 8)
        ' yield execution so Excel remains responsive and user can hit Esc
        DoEvents
        i = i + 1
    Loop
    
    ' Validate recordsToValidate
    i = 1
    Dim key As Variant
    For Each key In recordsToValidate.Keys
        Dim recordToValidate As RecordTuple
        Set recordToValidate = recordsToValidate.Item(key)
        Dim gburgAddress As Scripting.Dictionary
        Set gburgAddress = Lookup.gburgQuery(recordToValidate.GburgFormatRawAddress.Item(AddressKey.Full))
        
        recordToValidate.SetValidAddress gburgAddress
        
        If gburgAddress.Item(AddressKey.Full) <> vbNullString Then
            ' Valid address
            recordToValidate.SetInCity InCityCode.ValidInCity
            addresses.Add recordToValidate.key, recordToValidate
            ' NOTE choosing not to add to autocorrected since raw address was enough to match
            ' autocorrected is used to save Google lookups, Gaithersburg lookups are free
            ' However, Gaithersburg lookup can change zipcode, format of address, addition of Apt, etc.
        Else
            recordToValidate.SetInCity InCityCode.NotYetAutocorrected
            needsAutocorrect.Add recordToValidate.key, recordToValidate
        End If
        Application.StatusBar = "Validating record " & i & " of " & (UBound(recordsToValidate.Keys) + 1)
        i = i + 1
        DoEvents
    Next key
    

    Application.StatusBar = "Writing addresses and computing totals"
    
    writeAddressesComputeTotals addresses, needsAutocorrect, discards, autocorrected

    Application.StatusBar = appStatus
End Sub


