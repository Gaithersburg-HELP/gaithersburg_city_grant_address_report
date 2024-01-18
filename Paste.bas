Attribute VB_Name = "Paste"
Option Explicit

'@Folder("City_Grant_Address_Report.src")
'@EntryPoint
Public Sub PasteRecords()
    ActiveWorkbook.Worksheets.[_Default]("Interface").Activate
    Application.ScreenUpdating = False
    
    ' Find last data row (last row containing data)
    ActiveSheet.Range("A8").Select
    If Trim$(ActiveCell.Offset(1, 0).Value) <> vbNullString Then
        Selection.End(xlDown).Select
    End If
    ' Select next row after last data row
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.PasteSpecial (xlPasteValues)

    Application.ScreenUpdating = True
End Sub
