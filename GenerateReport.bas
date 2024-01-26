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
    Dim addresses As Scripting.Dictionary
    Set addresses = Records.loadAddresses("Addresses")
    
    Dim address As Variant
    For Each address In addresses.Keys()
        Dim record As RecordTuple
        Set record = addresses.Item(address)
        ActiveWorkbook.Sheets.[_Default]("Final Report").Cells(2, 1).Value = record.CleanInitials
        ActiveWorkbook.Sheets.[_Default]("Final Report").Cells(2, 2).Value = record.GburgFormatValidAddress.Item(AddressKey.StreetName)
    Next address
    
    ' TODO Load addresses
    ' Filter by valid in city addresses
    ' Get Gburg Format address: Odend'hal, O'neill
    ' City should be Gaithersburg
    ' State should be UCase(State)
    
    ' Loop through all addresses, print by offset
    ' Add "x" for quarters
    ' Worksheets("Final Report").Range("A2").Offset(ReportRowOffset, 10).Value = "x"
    
    ' Switch to the output worksheet and sort by Street Name,
    ' Street Number, Street Type and Apt Number
    
    
    'getBlankRow
    ' ActiveWorkbook.Sheets.[_Default]("Final Report").Range("A2:O2").Select
    'ActiveSheet.Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.Sort.SortFields.Clear
    With ActiveSheet.Sort
        .SortFields.Clear
        .SortFields.Add key:=Selection.Columns(3), Order:=xlAscending
        .SortFields.Add key:=Selection.Columns(2), Order:=xlAscending
        .SortFields.Add key:=Selection.Columns(4), Order:=xlAscending
        .SortFields.Add key:=Selection.Columns(6), Order:=xlAscending
        .Header = xlNo
        .SetRange Selection
        .Apply
    End With
    
    ActiveSheet.Range("A2").Select
End Sub
