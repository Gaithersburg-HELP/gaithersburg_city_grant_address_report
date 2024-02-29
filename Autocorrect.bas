Attribute VB_Name = "Autocorrect"
'@Folder("City_Grant_Address_Report.src")
Option Explicit

Public Const requestLimit As Long = 8000

Public Function getRemainingRequests() As Long
    Dim text As String
    text = ThisWorkbook.Sheets.[_Default]("Needs Autocorrect").Shapes("API Limit").TextFrame.Characters.text
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
    ThisWorkbook.Sheets.[_Default]("Needs Autocorrect").Shapes("API Limit").TextFrame.Characters.text = _
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
        
        Dim bypassUserVerifiedCheck As Boolean
        bypassUserVerifiedCheck = False
        
        ' check if user moved ValidNotInCity record to Needs Autocorrect and wants to reverify
        If recordToAutocorrect.UserVerified = True And _
           recordToAutocorrect.InCity = ValidNotInCity Then
            recordToAutocorrect.SetInCity InCityCode.NotYetAutocorrected
            bypassUserVerifiedCheck = True
        End If
        
        
        If ((recordToAutocorrect.UserVerified = False) Or (bypassUserVerifiedCheck)) And _
           usedRequests < getRemainingRequests() And _
           recordToAutocorrect.InCity = InCityCode.NotYetAutocorrected Then
           
            Dim formattedRawAddress As Scripting.Dictionary
            Set formattedRawAddress = recordToAutocorrect.GburgFormatRawAddress
            
            Dim validatedAddress As Scripting.Dictionary
            Set validatedAddress = Lookup.googleValidateQuery(formattedRawAddress.Item(addressKey.Full), _
                                                              recordToAutocorrect.RawCity, _
                                                              recordToAutocorrect.RawState, _
                                                              recordToAutocorrect.RawZip, addressAPIKey)
            
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
        Set gburgAddress = Lookup.gburgQuery(recordToAutocorrect.GburgFormatValidAddress.Item(addressKey.Full))
        
        If (gburgAddress.Item(addressKey.Full) <> vbNullString) Then
            ' NOTE addresses such as 600 S Frederick Ave which are in Gaithersburg database but
            ' NOT DPV deliverable will still be marked as valid
            
            ' in theory this should be the same as Google's valid address, but gburgQuery could return different zip
            recordToAutocorrect.SetValidAddress gburgAddress
            
            ' Addresses with unit will always match even if raw unit is incorrect
            ' because Gaithersburg has the same address without unit in their database
            ' Check for this and fail autocorrection if dropped raw unit
            If recordToAutocorrect.validUnitWithNum = vbNullString And _
               recordToAutocorrect.RawUnitWithNum <> vbNullString Then
                recordToAutocorrect.SetInCity InCityCode.FailedAutocorrectInCity
            Else
                recordToAutocorrect.SetInCity InCityCode.ValidInCity
                
                addresses.Add recordToAutocorrect.key, recordToAutocorrect
                addressesToAutocorrect.Remove recordToAutocorrect.key
                ' TODO rewrite code so that record has flag for which dictionaries it belongs in
                If recordToAutocorrect.isAutocorrected Then
                    autocorrected.Add recordToAutocorrect.key, recordToAutocorrect
                End If
            End If
        ElseIf isDPVConfirmed Then
            ' Gaithersburg database does not match USPS database on multiple addresses such as:
            ' - 110-150 Chevy Chase St Unit 102 > should be Apt 102
            ' - 25 Chestnut St Unit A > should be Ste A
            ' so double check by searching without unit
            Set gburgAddress = Lookup.gburgQuery(recordToAutocorrect.GburgFormatValidAddress.Item(addressKey.streetAddress))
            If gburgAddress.Item(addressKey.Full) <> vbNullString Then
                recordToAutocorrect.SetValidAddress gburgAddress
                recordToAutocorrect.SetInCity InCityCode.FailedAutocorrectInCity
            Else
                recordToAutocorrect.SetInCity InCityCode.ValidNotInCity
            
                addresses.Add recordToAutocorrect.key, recordToAutocorrect
                addressesToAutocorrect.Remove recordToAutocorrect.key
                ' TODO rewrite code so that record has flag for which dictionaries it belongs in
                If recordToAutocorrect.isAutocorrected Then
                    autocorrected.Add recordToAutocorrect.key, recordToAutocorrect
                End If
            End If
        ElseIf receivedValidation Then
            If Lookup.possibleInGburgQuery(minLongitude, minLatitude, maxLongitude, maxLatitude) Then
                recordToAutocorrect.SetInCity InCityCode.FailedAutocorrectInCity
            Else
                recordToAutocorrect.SetInCity InCityCode.FailedAutocorrectNotInCity
                
                addressesToAutocorrect.Remove recordToAutocorrect.key
                discards.Add recordToAutocorrect.key, recordToAutocorrect
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


