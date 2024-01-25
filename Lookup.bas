Attribute VB_Name = "Lookup"
'@Folder("City_Grant_Address_Report.src")
Option Explicit

' This macro subroutine (Ctrl+L) may be used to double-check
' those street addresses not flagged In-City by lookup on the
' Gaithersburg city address search page in browser window.
Public Sub LookupInCity() ' Ctrl+L
    Dim currentRow As String
    Dim GburgCityURL As String
    Dim AddrLookupURL As String
'    GburgCityURL = "https://maps.gaithersburgmd.gov/AddressSearch/index.html?address="
    GburgCityURL = "http://maps.gaithersburgmd.gov/AddressSearch/index.html?address="
' TODO address lookup by sheet
' Pick up Street Number, Street Name, and Street Type trimming outer spaces just in case
    'CurrentRow = ActiveCell.Row ' Get values from current row columns D and E
    'StreetNumber = Trim(Range("D" & CurrentRow).Value)
    'StreetName = Trim(Range("E" & CurrentRow).Value)
    'StreetType = Trim(Range("F" & CurrentRow).Value)
' Build the full lookup URL and replace inner spaces with plus signs
    'AddrLookupURL = GburgCityURL & StreetNumber & "+" & StreetName & "+" & StreetType
    AddrLookupURL = vbNullString
    AddrLookupURL = Replace(AddrLookupURL, " ", "+")
' Go to the Gaithersburg City Address Search site and lookup
    ActiveWorkbook.FollowHyperlink address:=AddrLookupURL
End Sub

Public Sub OpenAddressValidationWebsite()
    ActiveWorkbook.FollowHyperlink address:="https://developers.google.com/maps/documentation/address-validation/demo"
End Sub

Public Sub OpenUSPSZipcodeWebsite()
    ActiveWorkbook.FollowHyperlink address:="https://tools.usps.com/zip-code-lookup.htm?byaddress"
End Sub

Public Sub OpenGoogleMapsWebsite()
    ActiveWorkbook.FollowHyperlink address:="https://www.google.com/maps"
End Sub
