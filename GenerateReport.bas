Attribute VB_Name = "GenerateReport"
Option Explicit

'@Folder("City_Grant_Address_Report.src")
'@EntryPoint
Public Sub confirmGenerateFinalReport()
    Dim confirmResponse As VbMsgBoxResult
    confirmResponse = MsgBox("Are you sure you wish to generate the final report?", vbYesNo + vbQuestion, "Confirmation")
    If confirmResponse = vbNo Then
        Exit Sub
    End If
    
    generateFinalReport
End Sub

Public Sub generateFinalReport()

End Sub
