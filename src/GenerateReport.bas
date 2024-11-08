Attribute VB_Name = "GenerateReport"
'@Folder("City_Grant_Address_Report.src")
Option Explicit

Private Sub writeNonRxReportRecord(ByVal record As RecordTuple)
    Dim row As Range
    
    Set row = SheetUtilities.getBlankRow(NonRxReportSheet.name)
    
    row.Cells.Item(1, 1) = "Gaithersburg HELP"
    row.Cells.Item(1, 2) = record.GburgFormatValidAddress.Item(addressKey.streetNum)
    row.Cells.Item(1, 3) = record.GburgFormatValidAddress.Item(addressKey.PrefixedStreetname)
    row.Cells.Item(1, 4) = record.GburgFormatValidAddress.Item(addressKey.StreetType)
    row.Cells.Item(1, 5) = record.GburgFormatValidAddress.Item(addressKey.unitType)
    row.Cells.Item(1, 6) = record.GburgFormatValidAddress.Item(addressKey.unitNum)
    row.Cells.Item(1, 7) = "Gaithersburg"
    row.Cells.Item(1, 8) = "MD"
    row.Cells.Item(1, 9) = record.CleanInitials
    row.Cells.Item(1, 10) = record.householdTotal
    row.Cells.Item(1, 11) = record.eighteenPlusTotal
    row.Cells.Item(1, 12) = record.zeroToOneTotal + record.twoToSeventeenTotal
    
    Dim Quarters() As Boolean
    Quarters = record.Quarters
    If Quarters(1) Then row.Cells.Item(1, 13) = "x"
    If Quarters(2) Then row.Cells.Item(1, 14) = "x"
    If Quarters(3) Then row.Cells.Item(1, 15) = "x"
    If Quarters(4) Then row.Cells.Item(1, 16) = "x"
End Sub

Public Sub generateNonRxReport()
    SheetUtilities.getNonRxReportRng.Clear
    
    Dim addresses As Scripting.Dictionary
    Set addresses = records.loadAddresses(AddressesSheet.name)
    
    Dim address As Variant
    For Each address In addresses.Keys()
        Dim record As RecordTuple
        Set record = addresses.Item(address)
        
        If record.InCity = ValidInCity And record.visitData.count > 0 Then writeNonRxReportRecord record
    Next address
    
    SheetUtilities.SortSheet NonRxReportSheet.name
    
    ActiveSheet.Range("A2").Select
End Sub

Public Sub generateRxReport(ByVal records As RxRecords)
    ' TODO
End Sub
