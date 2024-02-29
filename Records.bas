Attribute VB_Name = "Records"
Option Explicit

'@Folder "City_Grant_Address_Report.src"
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
    
    record.AddVisit recordRowFirstCell.value, recordRowFirstCell.offset(0, 1).value
    record.UserVerified = False

    record.guestID = recordRowFirstCell.offset(0, 2).value
    record.FirstName = recordRowFirstCell.offset(0, 3).value
    record.LastName = recordRowFirstCell.offset(0, 4).value
    record.RawAddress = recordRowFirstCell.offset(0, 5).value
    record.RawUnitWithNum = recordRowFirstCell.offset(0, 6).value
    record.RawCity = recordRowFirstCell.offset(0, 7).value
    record.RawState = recordRowFirstCell.offset(0, 8).value
    record.RawZip = recordRowFirstCell.offset(0, 9).value
    record.householdTotal = recordRowFirstCell.offset(0, 10).value
    
    Dim rx As Double
    rx = recordRowFirstCell.offset(0, 11).value
    If rx <> 0 Then record.addRx recordRowFirstCell.value, rx
    
    Set loadRecordFromRaw = record
End Function

Public Function loadRecordFromSheet(ByVal recordRowFirstCell As Range) As RecordTuple
    Dim record As RecordTuple
    Set record = New RecordTuple
    
    Dim services() As String
    services = loadServiceNames(recordRowFirstCell.Worksheet.Name)
    
    record.SetInCity recordRowFirstCell.offset(0, 0).value
    record.UserVerified = CBool(recordRowFirstCell.offset(0, 1).value)
    record.validAddress = recordRowFirstCell.offset(0, 2).value
    record.validUnitWithNum = recordRowFirstCell.offset(0, 3).value
    record.ValidZipcode = recordRowFirstCell.offset(0, 4).value
    record.RawAddress = recordRowFirstCell.offset(0, 5).value
    record.RawUnitWithNum = recordRowFirstCell.offset(0, 6).value
    record.RawCity = recordRowFirstCell.offset(0, 7).value
    record.RawState = recordRowFirstCell.offset(0, 8).value
    record.RawZip = recordRowFirstCell.offset(0, 9).value
    record.guestID = recordRowFirstCell.offset(0, 10).value
    record.FirstName = recordRowFirstCell.offset(0, 11).value
    record.LastName = recordRowFirstCell.offset(0, 12).value
    record.householdTotal = recordRowFirstCell.offset(0, 13).value
    Set record.rxTotal = JsonConverter.ParseJson(recordRowFirstCell.offset(0, 14).value)
    
    Dim visitData As Scripting.Dictionary
    Set visitData = New Scripting.Dictionary
    
    Dim j As Long
    j = 1
    Do While j <= UBound(services) + 1
        Dim visitJson As String
        visitJson = recordRowFirstCell.offset(0, 14 + j).value
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
    Set sheet = ThisWorkbook.Worksheets.[_Default](sheetName)
       
    If sheet.Range("A2").value = vbNullString Then
        Set loadAddresses = addresses
        Exit Function
    End If
    
    Dim i As Long
    i = 2
    Do While i < getBlankRow(sheetName).row
        Dim recordRowFirstCell As Range
        Set recordRowFirstCell = sheet.rows.Item(i).Cells.Item(1, 1)
        
        Dim record As RecordTuple
        Set record = loadRecordFromSheet(recordRowFirstCell)
        
        addresses.Add record.key, record
        i = i + 1
    Loop

    Set loadAddresses = addresses
End Function

Public Sub writeAddress(ByVal sheetName As String, ByVal record As RecordTuple)
    Dim sheet As Worksheet
    Set sheet = ThisWorkbook.Worksheets.[_Default](sheetName)
    
    ' Saves column numbers per existing service
    Dim serviceCols As Scripting.Dictionary
    Set serviceCols = New Scripting.Dictionary
    
    If sheet.Range("A2").value <> vbNullString Then
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
    
    recordRow.Cells.Item(1, 1).value = record.InCityStr
    recordRow.Cells.Item(1, 2).value = record.UserVerified
    recordRow.Cells.Item(1, 3).value = record.validAddress
    recordRow.Cells.Item(1, 4).value = record.validUnitWithNum
    recordRow.Cells.Item(1, 5).value = record.ValidZipcode
    recordRow.Cells.Item(1, 6).value = record.RawAddress
    recordRow.Cells.Item(1, 7).value = record.RawUnitWithNum
    recordRow.Cells.Item(1, 8).value = record.RawCity
    recordRow.Cells.Item(1, 9).value = record.RawState
    recordRow.Cells.Item(1, 10).value = record.RawZip
    recordRow.Cells.Item(1, 11).value = record.guestID
    recordRow.Cells.Item(1, 12).value = record.FirstName
    recordRow.Cells.Item(1, 13).value = record.LastName
    recordRow.Cells.Item(1, 14).value = record.householdTotal
    
    recordRow.Cells.Item(1, 15).value = JsonConverter.ConvertToJson(record.rxTotal)
    
    Dim serviceToAdd As Variant
    For Each serviceToAdd In record.visitData.Keys
        Dim visitDataToAdd As String
        visitDataToAdd = JsonConverter.ConvertToJson(record.visitData.Item(serviceToAdd))
        
        If Not serviceCols.Exists(serviceToAdd) Then
            Dim newServiceCol As Long
            newServiceCol = 15 + 1 + UBound(serviceCols.Keys) + 1
            serviceCols.Add serviceToAdd, newServiceCol
            ThisWorkbook.Worksheets.[_Default](sheetName).Cells(1, newServiceCol).value = serviceToAdd
        End If
        
        recordRow.Cells.Item(1, serviceCols.Item(serviceToAdd)).value = visitDataToAdd
    Next serviceToAdd
End Sub

Public Sub writeAddresses(ByVal sheetName As String, ByVal addresses As Scripting.Dictionary)
    ClearSheet sheetName
    Dim key As Variant
    For Each key In addresses.Keys
        writeAddress sheetName, addresses.Item(key)
    Next key
End Sub

Private Sub incrementCountyTotal(ByVal record As RecordTuple)
    Dim monthVisited(1 To 12) As Boolean
    Dim uniqueGuestIDTotal(1 To 12) As Long
    Dim uniqueGuestIDHouseholdTotal(1 To 12) As Long
    Dim guestIDTotal(1 To 12) As Long
    Dim householdTotal(1 To 12) As Long
    
    Dim monthNum As Long
    
    Dim service As Variant
    For Each service In record.visitData.Keys
        Dim quarter As Variant
        For Each quarter In record.visitData.Item(service).Keys
            Dim visit As Variant
            For Each visit In record.visitData.Item(service).Item(quarter)
                monthNum = month(visit)
                guestIDTotal(monthNum) = guestIDTotal(monthNum) + 1
                householdTotal(monthNum) = householdTotal(monthNum) + record.householdTotal
                If Not monthVisited(monthNum) Then
                    uniqueGuestIDTotal(monthNum) = uniqueGuestIDTotal(monthNum) + 1
                    uniqueGuestIDHouseholdTotal(monthNum) = uniqueGuestIDHouseholdTotal(monthNum) + _
                                                            record.householdTotal
                    monthVisited(monthNum) = True
                End If
            Next visit
        Next quarter
    Next service
    
    Dim zip As String
    If record.ValidZipcode <> vbNullString Then
        zip = record.ValidZipcode
    Else
        zip = record.RawZip
    End If
    
    Dim uniqueCols As Scripting.Dictionary
    Set uniqueCols = SheetUtilities.uniqueCountyZipCols()
    
    Dim zipCol As Long
    
    If uniqueCols.Exists(zip) Then
        zipCol = uniqueCols.Item(zip)
    Else
        ' BUG assumes poorer city
    Select Case zip
        Case 20906
            If record.RawCity = "Ashton" Or record.RawCity = "Aspen Hill" Then
                zipCol = CountyTotalCols.zip20906AshtonAspenHill
            Else
                zipCol = CountyTotalCols.zip20906SilverSpring
            End If
        Case 20916
            If record.RawCity = "Ashton" Or record.RawCity = "Aspen Hill" Then
                zipCol = CountyTotalCols.zip20916AshtonAspenHill
            Else
                zipCol = CountyTotalCols.zip20916SilverSpring
            End If
        Case 20815
            If record.RawCity = "Chevy Chase" Or record.RawCity = "Clarksburg" Then
                zipCol = CountyTotalCols.zip20815ChevyChaseClarksburg
            Else
                zipCol = CountyTotalCols.zip20815Bethesda
            End If
        Case 20825
            If record.RawCity = "Chevy Chase" Or record.RawCity = "Clarksburg" Then
                zipCol = CountyTotalCols.zip20825ChevyChaseClarksburg
            Else
                zipCol = CountyTotalCols.zip20825Bethesda
            End If
        Case 20852
            If record.RawCity = "Bethesda" Then
                zipCol = CountyTotalCols.zip20852Bethesda
            Else
                zipCol = CountyTotalCols.zip20852Rockville
            End If
        Case 20904
            If record.RawCity = "Colesville" Or record.RawCity = "Damascus" Then
                zipCol = CountyTotalCols.zip20904ColesvilleDamascus
            Else
                zipCol = CountyTotalCols.zip20904SilverSpring
            End If
        Case 20905
            If record.RawCity = "Colesville" Or record.RawCity = "Damascus" Then
                zipCol = CountyTotalCols.zip20905ColesvilleDamascus
            Else
                zipCol = CountyTotalCols.zip20905SilverSpring
            End If
        Case 20914
            If record.RawCity = "Colesville" Or record.RawCity = "Damascus" Then
                zipCol = CountyTotalCols.zip20914ColesvilleDamascus
            Else
                zipCol = CountyTotalCols.zip20914SilverSpring
            End If
        Case 20874
            If record.RawCity = "Darnestown" Or record.RawCity = "Derwood" Or record.RawCity = "Dickerson" Then
                zipCol = CountyTotalCols.zip20874DarnestownDerwoodDickerson
            Else
                zipCol = CountyTotalCols.zip20874GarrettParkGermantownGlenEcho
            End If
        Case 20878
            If record.RawCity = "Darnestown" Or record.RawCity = "Derwood" Or record.RawCity = "Dickerson" Then
                zipCol = CountyTotalCols.zip20878DarnestownDerwoodDickerson
            ElseIf record.RawCity = "Poolesville" Or record.RawCity = "Potomac" Then
                zipCol = CountyTotalCols.zip20878PoolesvillePotomac
            Else
                zipCol = CountyTotalCols.zip20878Gaithersburg
            End If
        Case 20855
            If record.RawCity = "Darnestown" Or record.RawCity = "Derwood" Or record.RawCity = "Dickerson" Then
                zipCol = CountyTotalCols.zip20855DarnestownDerwoodDickerson
            Else
                zipCol = CountyTotalCols.zip20855Rockville
            End If
        Case 20877
            If record.RawCity = "Montgomery Village" Or record.RawCity = "Olney" Then
                zipCol = CountyTotalCols.zip20877MontgomeryVillageOlney
            Else
                zipCol = CountyTotalCols.zip20877Gaithersburg
            End If
        Case 20882
            If record.RawCity = "Kensington" Or record.RawCity = "Laytonsville" Then
                zipCol = CountyTotalCols.zip20882KensingtonLaytonsville
            Else
                zipCol = CountyTotalCols.zip20882Gaithersburg
            End If
        Case 20886
            If record.RawCity = "Montgomery Village" Or record.RawCity = "Olney" Then
                zipCol = CountyTotalCols.zip20886MontgomeryVillageOlney
            Else
                zipCol = CountyTotalCols.zip20886Gaithersburg
            End If
        Case 20879
            If record.RawCity = "Kensington" Or record.RawCity = "Laytonsville" Then
                zipCol = CountyTotalCols.zip20879KensingtonLaytonsville
            ElseIf record.RawCity = "Montgomery Village" Or record.RawCity = "Olney" Then
                zipCol = CountyTotalCols.zip20879MontgomeryVillageOlney
            Else
                zipCol = CountyTotalCols.zip20879Gaithersburg
            End If
        Case 20854
            If record.RawCity = "Poolesville" Or record.RawCity = "Potomac" Then
                zipCol = CountyTotalCols.zip20854PoolesvillePotomac
            Else
                zipCol = CountyTotalCols.zip20854Rockville
            End If
        Case 20859
            If record.RawCity = "Poolesville" Or record.RawCity = "Potomac" Then
                zipCol = CountyTotalCols.zip20859PoolesvillePotomac
            Else
                zipCol = CountyTotalCols.zip20859Rockville
            End If
        Case 20912
            If record.RawCity = "Sandy Spring" Or record.RawCity = "Spencerville" Or record.RawCity = "Takoma Park" Then
                zipCol = CountyTotalCols.zip20912SandySpringSpencervilleTakomaPark
            Else
                zipCol = CountyTotalCols.zip20912SilverSpring
            End If
        Case 20913
            If record.RawCity = "Sandy Spring" Or record.RawCity = "Spencerville" Or record.RawCity = "Takoma Park" Then
                zipCol = CountyTotalCols.zip20913SandySpringSpencervilleTakomaPark
            Else
                zipCol = CountyTotalCols.zip20913SilverSpring
            End If
        Case 20902
            If record.RawCity = "Washington Grove" Or record.RawCity = "Wheaton" Then
                zipCol = CountyTotalCols.zip20902WashingtonGroveWheaton
            Else
                zipCol = CountyTotalCols.zip20902SilverSpring
            End If
        Case 20915
            If record.RawCity = "Washington Grove" Or record.RawCity = "Wheaton" Then
                zipCol = CountyTotalCols.zip20915WashingtonGroveWheaton
            Else
                zipCol = CountyTotalCols.zip20915SilverSpring
            End If
        Case Else
            zipCol = -1
        End Select
    End If
    
    Dim i As Long
    For i = 1 To 12
        ' TODO off-by-one error on column enum
        getCountyRng.Cells.Item(i, CountyTotalCols.householdDuplicate - 1) = _
            getCountyRng.Cells.Item(i, CountyTotalCols.householdDuplicate - 1) + guestIDTotal(i)
        getCountyRng.Cells.Item(i, CountyTotalCols.householdUnduplicate - 1) = _
            getCountyRng.Cells.Item(i, CountyTotalCols.householdUnduplicate - 1) + uniqueGuestIDTotal(i)
        getCountyRng.Cells.Item(i, CountyTotalCols.individualDuplicate - 1) = _
            getCountyRng.Cells.Item(i, CountyTotalCols.individualDuplicate - 1) + householdTotal(i)
        getCountyRng.Cells.Item(i, CountyTotalCols.individualUnduplicate - 1) = _
            getCountyRng.Cells.Item(i, CountyTotalCols.individualUnduplicate - 1) + uniqueGuestIDHouseholdTotal(i)
            
        getCountyRng.Cells.Item(i, CountyTotalCols.poundsFood - 1) = _
            getCountyRng.Cells.Item(i, CountyTotalCols.poundsFood - 1) + (householdTotal(i) * 8)
        
        If zipCol <> -1 Then
            getCountyRng.Cells.Item(i, zipCol - 1) = getCountyRng.Cells.Item(i, zipCol - 1) + guestIDTotal(i)
        End If
        
    Next i
End Sub

Public Sub computeCountyTotals()
    getCountyRng.value = 0
    
    Dim addresses As Scripting.Dictionary
    Set addresses = Records.loadAddresses("Addresses")
    Dim key As Variant
    For Each key In addresses.Keys
        incrementCountyTotal addresses.Item(key)
    Next key
    
    Dim autocorrect As Scripting.Dictionary
    Set autocorrect = Records.loadAddresses("Needs Autocorrect")
    
    For Each key In autocorrect.Keys
        incrementCountyTotal autocorrect.Item(key)
    Next key
    
    Dim discards As Scripting.Dictionary
    Set discards = Records.loadAddresses("Discards")
    
    For Each key In discards.Keys
        incrementCountyTotal discards.Item(key)
    Next key
End Sub

Public Sub computeTotals()
    SheetUtilities.getTotalsRng.Clear
    SheetUtilities.getCountyRng.Clear
    
    Dim addresses As Scripting.Dictionary
    Set addresses = Records.loadAddresses("Addresses")

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
        
        ' Gaithersburg totals
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
End Sub

Public Sub writeAddressesComputeTotals(ByVal addresses As Scripting.Dictionary, _
                                       ByVal needsAutocorrect As Scripting.Dictionary, _
                                       ByVal discards As Scripting.Dictionary, _
                                       ByVal autocorrected As Scripting.Dictionary)
    SheetUtilities.ClearAll
       
    writeAddresses "Addresses", addresses
    writeAddresses "Needs Autocorrect", needsAutocorrect
    writeAddresses "Discards", discards
    writeAddresses "Autocorrected", autocorrected
    
    SortAll
    
    computeTotals
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
    i = getPastedRecordsRng.row
    Do While i < getBlankRow("Interface").row
        Dim recordToAdd As RecordTuple
        Set recordToAdd = loadRecordFromRaw(ThisWorkbook.Sheets.[_Default]("Interface").Range("A" & i))
        
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
            If autocorrected.Exists(recordToAdd.key) Then
                Set existingRecord = autocorrected.Item(recordToAdd.key)
                existingRecord.MergeRecord recordToAdd
            End If
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
        Set gburgAddress = Lookup.gburgQuery(recordToValidate.GburgFormatRawAddress.Item(addressKey.Full))
        
        recordToValidate.SetValidAddress gburgAddress
        
        If gburgAddress.Item(addressKey.Full) <> vbNullString Then
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
    computeCountyTotals
    
    Application.StatusBar = appStatus
End Sub


