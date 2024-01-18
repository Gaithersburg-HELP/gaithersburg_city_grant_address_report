Attribute VB_Name = "DeleteAllVisitData"
Option Explicit

'@Folder("City_Grant_Address_Report.src")
Public Sub DeleteAllVisitData()
    Dim confirmResponse As VbMsgBoxResult
    confirmResponse = MsgBox("Are you sure you wish to delete all visit data?", vbYesNo + vbQuestion, "Confirmation")
    If confirmResponse = vbNo Then
        Exit Sub
    End If
End Sub
