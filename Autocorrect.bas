Attribute VB_Name = "Autocorrect"
'@Folder("City_Grant_Address_Report.src")
Option Explicit

Public Const requestLimit As Long = 8000

Public Function getRemainingRequests() As Long
    Dim text As String
    text = ActiveWorkbook.Sheets.[_Default]("Needs Autocorrect").Shapes("API Limit").TextFrame.Characters.text
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
    ActiveWorkbook.Sheets.[_Default]("Needs Autocorrect").Shapes("API Limit").TextFrame.Characters.text = _
        num & " / " & requestLimit & " left until " & MonthName(month(Date) + 1)
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
    
    attemptValidation
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
        
        If recordToAutocorrect.UserVerified = False And _
           usedRequests < getRemainingRequests() And _
           recordToAutocorrect.InCity = InCityCode.NotYetAutocorrected Then
           
            Dim formattedRawAddress As Scripting.Dictionary
            Set formattedRawAddress = recordToAutocorrect.GburgFormatRawAddress
            
            Dim validatedAddress As Scripting.Dictionary
            Set validatedAddress = Lookup.googleValidateQuery(formattedRawAddress.Item(AddressKey.Full), _
                                                              recordToAutocorrect.RawCity, _
                                                              recordToAutocorrect.RawState, _
                                                              recordToAutocorrect.RawZip, addressAPIKey)
            
            If Not (validatedAddress Is Nothing) Then
                recordToAutocorrect.SetValidAddress validatedAddress
                minLongitude = validatedAddress.Item(AddressKey.minLongitude)
                maxLongitude = validatedAddress.Item(AddressKey.maxLongitude)
                minLatitude = validatedAddress.Item(AddressKey.minLatitude)
                maxLatitude = validatedAddress.Item(AddressKey.maxLatitude)
                
                receivedValidation = True
            
                If validatedAddress.Item(AddressKey.Full) <> vbNullString Then
                    isDPVConfirmed = True
                End If
            End If
            usedRequests = usedRequests + 1
        End If
                
        Dim gburgAddress As Scripting.Dictionary
        Set gburgAddress = Lookup.gburgQuery(recordToAutocorrect.GburgFormatValidAddress.Item(AddressKey.Full))
        
        If (gburgAddress.Item(AddressKey.Full) <> vbNullString) Then
            ' NOTE addresses such as 600 S Frederick Ave which are in Gaithersburg database but
            ' NOT DPV deliverable will still be marked as valid
            
            ' in theory this should be the same as Google's valid address, but gburgQuery could return different zip
            recordToAutocorrect.SetValidAddress gburgAddress
            recordToAutocorrect.SetInCity InCityCode.ValidInCity
            
            addresses.Add recordToAutocorrect.key, recordToAutocorrect
            addressesToAutocorrect.Remove recordToAutocorrect.key
            ' TODO rewrite code so that record has flag for which dictionaries it belongs in
            If recordToAutocorrect.isAutocorrected Then
                autocorrected.Add recordToAutocorrect.key, recordToAutocorrect
            End If
        ElseIf isDPVConfirmed Then
            ' Gaithersburg database does not match USPS database on multiple addresses such as:
            ' - 110-150 Chevy Chase St Unit 102 > should be Apt 102
            ' - 25 Chestnut St Unit A > should be Ste A
            ' so double check by searching without unit
            Set gburgAddress = Lookup.gburgQuery(recordToAutocorrect.GburgFormatValidAddress.Item(AddressKey.streetAddress))
            If gburgAddress.Item(AddressKey.Full) <> vbNullString Then
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
    
    printRemainingRequests (getRemainingRequests() - usedRequests)
    
    Application.StatusBar = appStatus
End Sub


