Attribute VB_Name = "DeleteVisitData"
Option Explicit

'@Folder("City_Grant_Address_Report.src")
'@EntryPoint
Public Sub DeleteAllVisitData()
    Dim confirmResponse As VbMsgBoxResult
    confirmResponse = MsgBox("Are you sure you wish to delete all visit data?", vbYesNo + vbQuestion, "Confirmation")
    If confirmResponse = vbNo Then
        Exit Sub
    End If
    
    getAddressVisitDataRng("Addresses").Clear
End Sub

'@EntryPoint
Public Sub DeleteService()
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
    
    'TODO discard all remaining
End Sub

'@EntryPoint
Public Sub confirmDiscardSelected()
    Dim confirmResponse As VbMsgBoxResult
    confirmResponse = MsgBox("Are you sure you wish to discard the selected record?", vbYesNo + vbQuestion, "Confirmation")
    If confirmResponse = vbNo Then
        Exit Sub
    End If
    
    'TODO discard selected
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
