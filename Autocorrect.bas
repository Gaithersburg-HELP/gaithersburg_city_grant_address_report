Attribute VB_Name = "Autocorrect"
'@Folder("City_Grant_Address_Report.src")
Option Explicit
Private Function getRemainingRequests() As Long
    'TODO get remaining requests, based on current time compared against refresh date
    getRemainingRequests = 1000
End Function
Private Sub printRemainingRequests(ByVal num As Long)
    'TODO print remaining requests and month refresh date
    ActiveSheet.Shapes("API Limit").TextFrame.Characters.Text = num & " / 8000 left"
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
    Dim addressAPIKey As String
    addressAPIKey = getAPIKeys().Item(addressValidationKeyname)
    
    Dim addresses As Scripting.Dictionary
    Set addresses = Records.loadAddresses("Needs Autocorrect")
    Debug.Print (addresses.Keys().Count & addressAPIKey)
    
    printRemainingRequests (1)
    ' TODO check if user has verified, if so then skip autocorrection and validate against Gaithersburg
    ' If fails Gaithersburg validation, remark as FALSE
    
    ' TODO autocorrecting
    ' autocorrectAddress(address)
    ' If autocorrected address is valid
        ' run against gaithersburg db
        ' Write to autocorrected addresses with json, highlight diff in yellow
        ' Add to address dictionary with gaithersburg result
    ' Else
        'add to discards dict, write to discards with autocorrect json
        ' If street name is in Gaithersburg street names
            ' highlight red
    
    ' autocorrectAddress
    ' Returns autocorrected address, address valid json, autocorrect json
    ' TODO write test for this function
    ' TODO Submit street name + Gaithersburg city only to place autocomplete
    ' ? Get list of street names from Gaithersburg, Autocorrect to closest street name
    ' Autocorrect Av to Ave, W Deer Pk to W Deer Park Rd
    ' Check postfixes
    'autocorrectAddress = Array(address, "valid json", "autocorrect json")
    
    'If usps cass valid, intersect polygon
    ' if in Gaithersburg, highlight for user verification(check 501 S Frederick Ave)
    ' Using valid address Google, use USPS standardized (test 2 nina ave (CT).
    ' usps returns odendhal and oneill, add apostrophe. postfix test: 775 kimberly ct e
    ' - Adding Gaithersburg does work: 15119 frederick rd, gaithersburg, md corrects to Rockville, MD (or possibly Woodbine, MD)
    ' - go with USPS CASS address comparison, google "replaced" or "spellCorrected" is not always present,
    ' - 501 frederick ave, gaithersburg, md replaces to 501 S Frederick Ave
    ' - 501 frederik returns "spellCorrected",
End Sub
