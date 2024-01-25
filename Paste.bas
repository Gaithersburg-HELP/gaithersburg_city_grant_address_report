Attribute VB_Name = "Paste"
Option Explicit

'@Folder("City_Grant_Address_Report.src")
'@EntryPoint
Public Sub PasteRecords()
    ActiveWorkbook.Worksheets.[_Default]("Interface").Activate
    Application.ScreenUpdating = False
    
    getBlankRow("Interface").Cells.Item(1, 1).Select
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.PasteSpecial (xlPasteValues)

    Application.ScreenUpdating = True
End Sub
