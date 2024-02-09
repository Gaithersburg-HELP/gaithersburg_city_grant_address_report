Attribute VB_Name = "Lookup"
'@Folder("City_Grant_Address_Report.src")
Option Explicit

Public Enum AddressKey
    StreetAddress = -1 ' StreetNum + PrefixedStreetName + StreetType + Postfix
    Full = 0 ' StreetAddress + UnitType + UnitNum
    StreetNum = 1
    PrefixedStreetName = 2 ' No postfix
    StreetType = 3
    Postfix = 4
    UnitType = 5
    UnitNum = 6
    unitWithNum = 7
    Zip = 8
    minLongitude = 9 ' Double
    maxLongitude = 10 ' Double
    minLatitude = 11 ' Double
    maxLatitude = 12 ' Double
End Enum

Private Function initAddressKey() As Scripting.Dictionary
    Dim address As Scripting.Dictionary
    Set address = New Scripting.Dictionary
    address.Add AddressKey.StreetAddress, vbNullString
    address.Add AddressKey.Full, vbNullString
    address.Add AddressKey.StreetNum, vbNullString
    address.Add AddressKey.PrefixedStreetName, vbNullString
    address.Add AddressKey.StreetType, vbNullString
    address.Add AddressKey.Postfix, vbNullString
    address.Add AddressKey.UnitType, vbNullString
    address.Add AddressKey.UnitNum, vbNullString
    address.Add AddressKey.unitWithNum, vbNullString
    address.Add AddressKey.Zip, vbNullString
    address.Add AddressKey.minLongitude, 0
    address.Add AddressKey.maxLongitude, 0
    address.Add AddressKey.minLatitude, 0
    address.Add AddressKey.maxLatitude, 0
    Set initAddressKey = address
End Function

' Executes query and returns JSON dictionary. Dictionary will be Nothing if there is an error
Private Function sendQuery(ByVal requestMethod As String, ByVal url As String, _
                           ByVal contentType As String, ByVal payload As String) As Scripting.Dictionary
    Dim service As Object
    Set service = New MSXML2.XMLHTTP60
    
    Dim queryResult As String
    
    With service
        .Open requestMethod, url, False
        .setRequestHeader "Content-Type", contentType
        .send payload
        
        Do While .readyState <> 4
            DoEvents
        Loop
        
        If .Status >= 400 And .Status <= 599 Then
            queryResult = CStr(.Status) + " - " + .statusText
            If .responseText <> vbNullString Then
                queryResult = queryResult + vbCrLf & .responseText
            End If
            MsgBox "Error " + queryResult, vbCritical, "Connection"
            Set sendQuery = Nothing
        ElseIf .responseText <> vbNullString Then
            Set sendQuery = JsonConverter.ParseJson(.responseText)
        Else
            Set sendQuery = Nothing
        End If
    End With
End Function

' Executes REST query on Gaithersburg ArcGIS website to see if address is in city or not
' Expects Gaithersburg full formatted address
' Returns address dictionary of validated fields. All keys will exist but may be set to vbNullString or 0
' - If all address fields validated, AddressKey.Full will not be vbNullString
' - Since only one feature will be returned if not searching with wildcard %, will always be a fully validated address
Public Function gburgQuery(ByVal fullAddress As String) As Scripting.Dictionary
    Dim validatedAddress As Scripting.Dictionary
    Set validatedAddress = initAddressKey()
    
    ' ' is escaped as ''
    Dim formatAddress As String
    formatAddress = Replace(fullAddress, "'", "''")

    
    Dim queryString As String
    queryString = "https://maps.gaithersburgmd.gov/arcgis/rest/services/layers/GaithersburgCityAddresses/MapServer/0/query?" & _
        "f=json&" & "returnGeometry=false&" & _
        "outFields=Full_Address,Address_Number,Road_Prefix_Dir,Road_Name,Road_Type,Road_Post_Dir,Unit_Type,Unit_Number,Zip_Code&" & _
        "where=Full_Address%20LIKE%20%27" & _
        WorksheetFunction.EncodeURL(formatAddress) & "%27"
           
    Dim jsonResult As Scripting.Dictionary
    Set jsonResult = sendQuery("GET", queryString, "application/x-www-form-urlencoded", vbNullString)
    
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
            validatedAddress.Item(AddressKey.unitWithNum) = validatedAddress.Item(AddressKey.UnitType) & " " & _
                                                            validatedAddress.Item(AddressKey.UnitNum)
            validatedAddress.Item(AddressKey.Full) = validatedAddress.Item(AddressKey.StreetAddress) & " " & _
                                                     validatedAddress.Item(AddressKey.unitWithNum)

        Else
            validatedAddress.Item(AddressKey.Full) = validatedAddress.Item(AddressKey.StreetAddress)
        End If
    End If

    Set gburgQuery = validatedAddress
End Function

' Executes REST query on Gaithersburg ARCGIS to see if address envelope
' intersects Gaithersburg boundaries
Public Function possibleInGburgQuery(ByVal minLongitude As Double, ByVal minLatitude As Double, _
                                     ByVal maxLongitude As Double, ByVal maxLatitude As Double) As Boolean
    Dim queryString As String
    queryString = "https://maps.gaithersburgmd.gov/arcgis/rest/services/layers/basicLayers/MapServer/17/query?" & _
        "f=json&" & "returnGeometry=false&" & _
        "outFields=OBJECTID&" & _
        "where=1%3D1&" & _
        "geometryType=esriGeometryEnvelope&" & "inSR=4326&" & "spatialRel=esriSpatialRelIntersects&" & _
        "geometry=" & minLongitude & "%2C" & minLatitude & "%2C" & maxLongitude & "%2C" & maxLatitude
    
    Dim jsonResult As Scripting.Dictionary
    Set jsonResult = sendQuery("GET", queryString, "application/x-www-form-urlencoded", vbNullString)
    
    Set possibleInGburgQuery = (Not (jsonResult Is Nothing)) And jsonResult.Item("features").Count > 0
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

