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

Public Sub addRecords()
    
End Sub
