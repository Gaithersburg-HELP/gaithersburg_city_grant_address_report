Attribute VB_Name = "PasteRecords"
Option Explicit

'@Folder("City_Grant_Address_Report.src")
Public Sub PasteRecords()
    ActiveWorkbook.Worksheets("Interface").Activate
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
