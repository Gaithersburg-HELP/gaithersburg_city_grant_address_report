Attribute VB_Name = "Autocorrect"
'@Folder("City_Grant_Address_Report.src")
Option Explicit
Private Function getRemainingRequests() As Long
    'TODO get remaining requests, based on current time compared against refresh date
    getRemainingRequests = 1000
End Function
Private Sub printRemainingRequests(ByVal num As Long)
    'TODO print remaining requests and month refresh date
    ActiveSheet.Shapes("API Limit").TextFrame.Characters.Text = "8000 / 8000 left"
End Sub

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

Public Sub confirmDiscardAll()
    Dim confirmResponse As VbMsgBoxResult
    confirmResponse = MsgBox("Are you sure you wish to discard all records?", vbYesNo + vbQuestion, "Confirmation")
    If confirmResponse = vbNo Then
        Exit Sub
    End If
    
    'TODO discard all remaining
End Sub

Public Sub attemptValidation()
    ' TODO check if user has verified, if so then skip autocorrection and validate against Gaithersburg
    
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
End Sub
