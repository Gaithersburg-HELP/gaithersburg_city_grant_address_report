Attribute VB_Name = "GenerateFinalReport"
Option Explicit

'@Folder("City_Grant_Address_Report.src")
Public Sub GenerateFinalReport()
    Dim confirmResponse As VbMsgBoxResult
    confirmResponse = MsgBox("Are you sure you wish to generate the final report?", vbYesNo + vbQuestion, "Confirmation")
    If confirmResponse = vbNo Then
        Exit Sub
    End If
    
    ActiveCell.Range("A10").Value = "hello there"
End Sub
