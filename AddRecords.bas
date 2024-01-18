Attribute VB_Name = "AddRecords"
Option Explicit

'@Folder "City_Grant_Address_Report.src"
Public Sub AddRecords()
    Dim confirmResponse As VbMsgBoxResult
    confirmResponse = MsgBox("Are you sure you wish to add records?", vbYesNo + vbQuestion, "Confirmation")
    If confirmResponse = vbNo Then
        Exit Sub
    End If
End Sub
