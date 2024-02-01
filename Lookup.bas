Attribute VB_Name = "Lookup"
'@Folder("City_Grant_Address_Report.src")
Option Explicit

Public Enum AddressKey
    StreetAddress = -1 ' StreetNum + PrefixedStreetName + StreetType + Postfix
    Full = 0 ' StreetAddress + UnitType + UnitNum
    StreetNum = 1
    PrefixedStreetName = 2
    StreetType = 3
    Postfix = 4
    UnitType = 5
    UnitNum = 6
    Zip = 7
End Enum

' Executes REST query on Gaithersburg ArcGIS website to see if address is in city or not
' Expects Gaithersburg full formatted address
' Returns address dictionary of validated fields. All keys will exist but may be set to vbNullString
' - If all fields validated, AddressKey.Full will not be vbNullString
' - Since only one feature will be returned if not searching with wildcard %, will always be a fully validated address
Public Function gburgQuery(ByVal fullAddress As String) As Scripting.Dictionary
    Dim validatedAddress As Scripting.Dictionary
    Set validatedAddress = New Scripting.Dictionary
    validatedAddress.Add AddressKey.StreetAddress, vbNullString
    validatedAddress.Add AddressKey.Full, vbNullString
    validatedAddress.Add AddressKey.StreetNum, vbNullString
    validatedAddress.Add AddressKey.PrefixedStreetName, vbNullString
    validatedAddress.Add AddressKey.StreetType, vbNullString
    validatedAddress.Add AddressKey.Postfix, vbNullString
    validatedAddress.Add AddressKey.UnitType, vbNullString
    validatedAddress.Add AddressKey.UnitNum, vbNullString
    validatedAddress.Add AddressKey.Zip, vbNullString
    
    ' ' is escaped as ''
    Dim formatAddress As String
    formatAddress = Replace(fullAddress, "'", "''")

    Dim service As Object
    Set service = New MSXML2.XMLHTTP60
    Dim queryString As String
    queryString = "https://maps.gaithersburgmd.gov/arcgis/rest/services/layers/GaithersburgCityAddresses/MapServer/0/query?" & _
        "f=json&" & "returnGeometry=false&" & _
        "outFields=Full_Address,Address_Number,Road_Prefix_Dir,Road_Name,Road_Type,Road_Post_Dir,Unit_Type,Unit_Number,Zip_Code&" & _
        "where=Full_Address%20LIKE%20%27" & _
        WorksheetFunction.EncodeURL(formatAddress) & "%27"
    
    Dim queryResult As String
    With service
        .Open "GET", queryString, False
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        .send
        Do While .readyState <> 4
            DoEvents
        Loop
        
        If .Status >= 400 And .Status <= 599 Then
            queryResult = CStr(.Status) + " - " + .statusText
            If .responseText <> vbNullString Then
                queryResult = queryResult + vbCrLf & .responseText
            End If
            MsgBox "Error " + queryResult, vbCritical, "Connection"
            Set gburgQuery = validatedAddress
            Exit Function
        End If
        
        ' Rubberduck inspection bug see https://github.com/rubberduck-vba/Rubberduck/issues/6142
        '@Ignore AssignmentNotUsed
        queryResult = .responseText
    End With
    
    Dim jsonResult As Scripting.Dictionary
    
    If queryResult <> vbNullString Then
        Set jsonResult = JsonConverter.ParseJson(queryResult)
        If (Not (jsonResult Is Nothing)) And jsonResult.Item("features").Count > 0 Then
            ' Since searching on Full_Address, expect only one feature to be returned
            Dim gburgAddress As Scripting.Dictionary
            Set gburgAddress = jsonResult.Item("features").Item(1).Item("attributes")
            
            validatedAddress.Item(AddressKey.StreetNum) = gburgAddress.Item("Address_Number")
            
            If gburgAddress.Item("Road_Prefix_Dir") <> vbNullString Then
                validatedAddress.Item(AddressKey.PrefixedStreetName) = gburgAddress.Item("Road_Prefix_Dir") & " " & _
                                                                       gburgAddress.Item("Road_Name")
            Else
                validatedAddress.Item(AddressKey.PrefixedStreetName) = gburgAddress.Item("Road_Name")
            End If
            
            validatedAddress.Item(AddressKey.StreetType) = gburgAddress.Item("Road_Type")
            validatedAddress.Item(AddressKey.Postfix) = gburgAddress.Item("Road_Post_Dir")
            validatedAddress.Item(AddressKey.UnitType) = gburgAddress.Item("Unit_Type")
            validatedAddress.Item(AddressKey.UnitNum) = gburgAddress.Item("Unit_Number")
            validatedAddress.Item(AddressKey.Zip) = gburgAddress.Item("Zip_Code")
            
            validatedAddress.Item(AddressKey.StreetAddress) = validatedAddress.Item(AddressKey.StreetNum) & " " & _
                                                              validatedAddress.Item(AddressKey.PrefixedStreetName) & " " & _
                                                              validatedAddress.Item(AddressKey.StreetType)
            If validatedAddress.Item(AddressKey.Postfix) <> vbNullString Then
                validatedAddress.Item(AddressKey.StreetAddress) = validatedAddress.Item(AddressKey.StreetAddress) & " " & _
                                                                  validatedAddress.Item(AddressKey.Postfix)
            End If
            
            If validatedAddress.Item(AddressKey.UnitType) <> vbNullString Then
                validatedAddress.Item(AddressKey.Full) = validatedAddress.Item(AddressKey.StreetAddress) & " " & _
                                                         validatedAddress.Item(AddressKey.UnitType) & " " & _
                                                         validatedAddress.Item(AddressKey.UnitNum)
            Else
                validatedAddress.Item(AddressKey.Full) = validatedAddress.Item(AddressKey.StreetAddress)
            End If
        End If
    End If
    Set gburgQuery = validatedAddress
End Function

' This macro subroutine may be used to double-check
' street addresses by lookup on the Gaithersburg city address search page in browser window.
'@EntryPoint
Public Sub LookupInCity()
    Dim currentRowFirstCell As Range
    Set currentRowFirstCell = ActiveWorkbook.ActiveSheet.Cells.Item(ActiveCell.row, 1)
    
    Dim record As RecordTuple
    Set record = Records.loadRecordFromSheet(currentRowFirstCell)
    
    Dim AddrLookupURL As String
    AddrLookupURL = "https://maps.gaithersburgmd.gov/AddressSearch/index.html?address="
    AddrLookupURL = AddrLookupURL & record.GburgFormatRawAddress.Item(AddressKey.StreetAddress)
    AddrLookupURL = Replace(AddrLookupURL, " ", "+")
    
    ActiveWorkbook.FollowHyperlink address:=AddrLookupURL
End Sub

'@EntryPoint
Public Sub OpenAddressValidationWebsite()
    ActiveWorkbook.FollowHyperlink address:="https://developers.google.com/maps/documentation/address-validation/demo"
End Sub

'@EntryPoint
Public Sub OpenUSPSZipcodeWebsite()
    ActiveWorkbook.FollowHyperlink address:="https://tools.usps.com/zip-code-lookup.htm?byaddress"
End Sub

'@EntryPoint
Public Sub OpenGoogleMapsWebsite()
    ActiveWorkbook.FollowHyperlink address:="https://www.google.com/maps"
End Sub

