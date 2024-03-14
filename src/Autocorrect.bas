Attribute VB_Name = "Autocorrect"
'@Folder("City_Grant_Address_Report.src")
Option Explicit

Public Const requestLimit As Long = 8000

Public Function getRemainingRequests() As Long
    Dim text As String
    text = AutocorrectAddressesSheet.Shapes.[_Default]("API Limit").TextFrame.Characters.text
    Dim refreshMonth As String
    refreshMonth = Lookup.RWordTrim(text)(1)
    If month(DateValue(refreshMonth & " 1 2024")) = month(Date) Then
        ' Limit refreshed this month
        printRemainingRequests (requestLimit)
        getRemainingRequests = requestLimit
    Else
        getRemainingRequests = CLng(Split(text, " ")(0))
    End If
End Function

Public Sub printRemainingRequests(ByVal num As Long)
    AutocorrectAddressesSheet.Shapes.[_Default]("API Limit").TextFrame.Characters.text = _
        num & " / " & requestLimit & " left until " & MonthName(month(Date) + 1)
End Sub

Public Sub attemptValidation()
    Dim appStatus As Variant
    If Application.StatusBar = False Then appStatus = False Else appStatus = Application.StatusBar
    
    Application.StatusBar = "Loading addresses"
    
    Dim addressAPIKey As String
    addressAPIKey = getAPIKeys().Item(addressValidationKeyname)
    
    Dim addresses As Scripting.Dictionary
    Set addresses = Records.loadAddresses("Addresses")
    
    Dim autocorrected As Scripting.Dictionary
    Set autocorrected = Records.loadAddresses("Autocorrected")
    
    Dim addressesToAutocorrect As Scripting.Dictionary
    Set addressesToAutocorrect = Records.loadAddresses("Needs Autocorrect")
    
    Dim discards As Scripting.Dictionary
    Set discards = Records.loadAddresses("Discards")
    
    If getRemainingRequests() < (UBound(addressesToAutocorrect.Keys) + 1) Then
        MsgBox "Insufficient requests remaining to autocorrect all addresses, " & _
               "not all addresses will be autocorrected.", vbOKOnly, "Error"
    End If
    
    Dim usedRequests As Long
    usedRequests = 0
    
    Dim i As Long
    i = 1
    
    Dim recordProgressLimit As Long
    recordProgressLimit = UBound(addressesToAutocorrect.Keys()) + 1
    
    Dim recordKey As Variant
    For Each recordKey In addressesToAutocorrect.Keys
        Dim recordToAutocorrect As RecordTuple
        Set recordToAutocorrect = addressesToAutocorrect.Item(recordKey)
        
        Dim isDPVConfirmed As Boolean
        isDPVConfirmed = False
        
        Dim receivedValidation As Boolean
        receivedValidation = False
        
        Dim minLongitude As Double
        Dim maxLongitude As Double
        Dim minLatitude As Double
        Dim maxLatitude As Double
        
        ' if user verified, only re-run google validation if fixing a ValidNotInCity address to verify if ValidNotInCity
        ' other user verified codes like FailedAutocorrection or ValidInCity should be verified in Gaithersburg database
        If (usedRequests < getRemainingRequests) And _
           ((recordToAutocorrect.UserVerified = True And recordToAutocorrect.InCity = InCityCode.ValidNotInCity) Or _
            (recordToAutocorrect.InCity = InCityCode.NotYetAutocorrected)) Then
            
            Dim validatedAddress As Scripting.Dictionary
            If recordToAutocorrect.UserVerified = True Then
                ' TODO if adding valid city field fix this
                Set validatedAddress = Lookup.googleValidateQuery( _
                                        recordToAutocorrect.GburgFormatValidAddress.Item(addressKey.Full), _
                                        recordToAutocorrect.RawCity, _
                                        recordToAutocorrect.RawState, _
                                        recordToAutocorrect.ValidZipcode, addressAPIKey)
            Else
                Set validatedAddress = Lookup.googleValidateQuery( _
                                        recordToAutocorrect.GburgFormatRawAddress.Item(addressKey.Full), _
                                        recordToAutocorrect.RawCity, _
                                        recordToAutocorrect.RawState, _
                                        recordToAutocorrect.RawZip, addressAPIKey)
            End If
            
            If Not (validatedAddress Is Nothing) Then
                recordToAutocorrect.SetValidAddress validatedAddress
                minLongitude = validatedAddress.Item(addressKey.minLongitude)
                maxLongitude = validatedAddress.Item(addressKey.maxLongitude)
                minLatitude = validatedAddress.Item(addressKey.minLatitude)
                maxLatitude = validatedAddress.Item(addressKey.maxLatitude)
                
                receivedValidation = True
            
                If validatedAddress.Item(addressKey.Full) <> vbNullString Then
                    isDPVConfirmed = True
                End If
            End If
            usedRequests = usedRequests + 1
        End If
                
        Dim gburgAddress As Scripting.Dictionary
        ' Note that this does NOT use Google's addressKey.Full, this builds GburgFormatValidAddress from base valid fields
        Set gburgAddress = Lookup.gburgQuery(recordToAutocorrect.GburgFormatValidAddress.Item(addressKey.Full))
        
        If (gburgAddress.Item(addressKey.Full) <> vbNullString) Then
            ' NOTE addresses such as 600 S Frederick Ave which are in Gaithersburg database but
            ' NOT DPV deliverable will still be marked as valid
            
            ' in theory this should be the same as Google's valid address, but gburgQuery could return different zip
            recordToAutocorrect.SetValidAddress gburgAddress
            
            Dim isSingleMatch As Boolean
            isSingleMatch = False
            
            ' Addresses with unit will always match even if raw unit is incorrect
            ' because Gaithersburg has the same address without unit in their database
            ' However, some addresses like 497 Quince Orchard Rd are Motel 6 and unit can be dropped safely
            ' because there's only one match in Gaithersburg database
            ' Check for this and fail autocorrection if dropped raw unit and more than one match
            ' However user verified records should be allowed because sometimes unit number is wrong and can't be corrected
            If recordToAutocorrect.UserVerified = False And _
               recordToAutocorrect.validUnitWithNum = vbNullString And _
               recordToAutocorrect.RawUnitWithNum <> vbNullString Then
                Dim count As Long
                count = Lookup.gburgPartialQuery(gburgAddress.Item(addressKey.Full))
                If count = 1 Then
                    isSingleMatch = True
                Else
                    recordToAutocorrect.SetInCity InCityCode.FailedAutocorrectInCity
                End If
            Else
                isSingleMatch = True
            End If
            If isSingleMatch Or recordToAutocorrect.UserVerified Then
                recordToAutocorrect.SetInCity InCityCode.ValidInCity
                
                addresses.Add recordToAutocorrect.key, recordToAutocorrect
                addressesToAutocorrect.Remove recordToAutocorrect.key
                ' TODO rewrite code so that record has flag for which dictionaries it belongs in
                If recordToAutocorrect.isAutocorrected Then
                    autocorrected.Add recordToAutocorrect.key, recordToAutocorrect
                End If
            End If
        ElseIf receivedValidation Then
            If Lookup.possibleInGburgQuery(minLongitude, minLatitude, maxLongitude, maxLatitude) Then
                ' Gaithersburg database does not match USPS database on multiple addresses such as:
                ' - 110-150 Chevy Chase St Unit 102 > should be Apt 102
                ' - 25 Chestnut St Unit A > should be Ste A
                ' - 319 N Summit Dr > should be Ave
                ' - USPS returns 738 Quince Orch instead of Quince Orchard for some reason
                recordToAutocorrect.SetInCity InCityCode.FailedAutocorrectInCity
            Else
                If isDPVConfirmed Then
                    recordToAutocorrect.SetInCity InCityCode.ValidNotInCity
                    addresses.Add recordToAutocorrect.key, recordToAutocorrect
                Else
                    recordToAutocorrect.SetInCity InCityCode.FailedAutocorrectNotInCity
                    discards.Add recordToAutocorrect.key, recordToAutocorrect
                End If
                
                addressesToAutocorrect.Remove recordToAutocorrect.key
                ' TODO rewrite code so that record has flag for which dictionaries it belongs in
                If recordToAutocorrect.isAutocorrected Then
                    autocorrected.Add recordToAutocorrect.key, recordToAutocorrect
                End If
            End If
        End If
        
        Application.StatusBar = "Validating record " & i & " of " & recordProgressLimit
        i = i + 1
        DoEvents
    Next recordKey
    
    Application.StatusBar = "Writing addresses"
    
    Records.writeAddressesComputeTotals addresses, addressesToAutocorrect, discards, autocorrected
    Records.computeCountyTotals
    
    printRemainingRequests (getRemainingRequests() - usedRequests)
    
    Application.StatusBar = appStatus
End Sub


