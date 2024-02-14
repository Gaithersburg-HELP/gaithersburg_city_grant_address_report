Attribute VB_Name = "Lookup"
'@Folder("City_Grant_Address_Report.src")
Option Explicit

Public Enum AddressKey
    streetAddress = -1 ' StreetNum + PrefixedStreetName + StreetType + Postfix
    Full = 0 ' StreetAddress + UnitType + UnitNum
    streetNum = 1
    PrefixedStreetName = 2 ' No postfix
    StreetType = 3
    Postfix = 4
    UnitType = 5
    UnitNum = 6
    unitWithNum = 7
    zip = 8
    minLongitude = 9 ' Double
    maxLongitude = 10 ' Double
    minLatitude = 11 ' Double
    maxLatitude = 12 ' Double
End Enum

Private Function initAddressKey() As Scripting.Dictionary
    Dim address As Scripting.Dictionary
    Set address = New Scripting.Dictionary
    address.Add AddressKey.streetAddress, vbNullString
    address.Add AddressKey.Full, vbNullString
    address.Add AddressKey.streetNum, vbNullString
    address.Add AddressKey.PrefixedStreetName, vbNullString
    address.Add AddressKey.StreetType, vbNullString
    address.Add AddressKey.Postfix, vbNullString
    address.Add AddressKey.UnitType, vbNullString
    address.Add AddressKey.UnitNum, vbNullString
    address.Add AddressKey.unitWithNum, vbNullString
    address.Add AddressKey.zip, vbNullString
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
            ' Debug.Print .responseText
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
' - Since only one feature will be returned if not searching with wildcard %,
'   will always be a fully validated address or vbNullString
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
        
        validatedAddress.Item(AddressKey.streetNum) = gburgAddress.Item("Address_Number")
        
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
        validatedAddress.Item(AddressKey.zip) = gburgAddress.Item("Zip_Code")
        
        validatedAddress.Item(AddressKey.streetAddress) = validatedAddress.Item(AddressKey.streetNum) & " " & _
                                                          validatedAddress.Item(AddressKey.PrefixedStreetName) & " " & _
                                                          validatedAddress.Item(AddressKey.StreetType)
        If validatedAddress.Item(AddressKey.Postfix) <> vbNullString Then
            validatedAddress.Item(AddressKey.streetAddress) = validatedAddress.Item(AddressKey.streetAddress) & " " & _
                                                              validatedAddress.Item(AddressKey.Postfix)
        End If
        
        If validatedAddress.Item(AddressKey.UnitType) <> vbNullString Then
            validatedAddress.Item(AddressKey.unitWithNum) = validatedAddress.Item(AddressKey.UnitType) & " " & _
                                                            validatedAddress.Item(AddressKey.UnitNum)
            validatedAddress.Item(AddressKey.Full) = validatedAddress.Item(AddressKey.streetAddress) & " " & _
                                                     validatedAddress.Item(AddressKey.unitWithNum)
        Else
            validatedAddress.Item(AddressKey.Full) = validatedAddress.Item(AddressKey.streetAddress)
        End If
    End If

    Set gburgQuery = validatedAddress
End Function

' Trims off last word including space before
' Returns [trimmed string, trimmed last word (blank if only one word)]
Public Function RWordTrim(ByVal str As String) As String()
    Dim lastWord As String
    Dim spaceIndex As Long
    spaceIndex = InStrRev(str, " ", -1, vbTextCompare)
    If (spaceIndex <> 0) Then
        lastWord = Right$(str, Len(str) - spaceIndex)
        RWordTrim = Split(Left$(str, spaceIndex - 1) & "|" & lastWord, "|")
    Else
        RWordTrim = Split(str & "|", "|")
    End If
End Function

' Executes REST query on Google Address Validation API to validate and autocorrect
' Returns address dictionary of AddressKey Full, StreetAddress, UnitWithNum, Zip, min/max lat/long
' - All other keys are set to vbNullString
' - If all fields validated, AddressKey.Full will not be vbNullString
'@Ignore AssignedByValParameter
Public Function googleValidateQuery(ByVal fullAddress As String, ByVal city As String, _
                                    ByVal state As String, ByVal zip As String, _
                                    ByVal apiKey As String) As Scripting.Dictionary
    Dim validatedAddress As Scripting.Dictionary
    Set validatedAddress = initAddressKey()

    Dim url As String
    url = "https://addressvalidation.googleapis.com/v1:validateAddress?key=" & apiKey

    ' Adding Gaithersburg does work: 15119 frederick rd, gaithersburg, md corrects to Rockville, MD (instead of Woodbine, MD)
    If city = vbNullString Then city = "Gaithersburg"
    If state = vbNullString Then state = "MD"
    If zip = vbNullString Then zip = "20878"

    ' using enableUspsCass returns inferior results
    ' see https://issuetracker.google.com/issues/325309557
    Dim payload As String
    payload = "{""address"": {""regionCode"":""US""," & _
                """locality"": """ & city & """," & _
                """administrativeArea"": """ & state & """," & _
                """postalCode"": """ & zip & """," & _
                """addressLines"": [""" & fullAddress & """]}}"
                
    Dim jsonResult As Scripting.Dictionary
    Set jsonResult = sendQuery("POST", url, "application/json", payload)

    If Not (jsonResult Is Nothing) Then
        On Error GoTo JsonError
        ' Go with USPS CASS address response first, cannot trust Google components
        ' E.g. 600 s frederik dr at 100, gburg, md, 20877, usa: Google components keeps "at 100" instead of Ste 100

        If jsonResult.Item("result").Exists("uspsData") Then
            Dim uspsFullAddress As String
            uspsFullAddress = jsonResult.Item("result").Item("uspsData").Item("standardizedAddress").Item("firstAddressLine")
            
            ' For DPV confirmation values, https://developers.google.com/maps/documentation/address-validation/handle-us-address
            ' Short circuiting does not exist in VBA
            If jsonResult.Item("result").Item("uspsData").Exists("dpvConfirmation") Then
                If jsonResult.Item("result").Item("uspsData").Item("dpvConfirmation") = "Y" Then
                    validatedAddress.Item(AddressKey.Full) = uspsFullAddress
                End If
            End If
            
            ' Cannot rely on hasInferredComponents to check if address changed (always true when ZIP extension is not provided)
            ' Cannot rely on hasReplacedComponents either, only true if e.g. city is replaced
            ' google "replaced" or "spellCorrected" is not always present
            '   i.e. "fredrick" doesn't show spellcorrected but "frederik" does
            
            validatedAddress.Item(AddressKey.zip) = jsonResult.Item("result").Item("uspsData") _
                                                              .Item("standardizedAddress").Item("zipCode")
            
            ' TODO? If not dpv confirmed, Get list of street names from Gaithersburg, Autocorrect to closest street name
            
            Dim streetAddress As String
            streetAddress = uspsFullAddress
            
            ' split on unit type if address input to unit level
            If jsonResult.Item("result").Item("verdict").Item("inputGranularity") = "SUB_PREMISE" Then
                ' see https://public-dhhs.ne.gov/nfocus/HowDoI/howdoi/usps_address_unit_types.htm
                
                Dim splitRWordArr() As String
                splitRWordArr = RWordTrim(uspsFullAddress)
                
                ' several designators such as upper don't have a secondary number
                ' however, as of 2/8/24 Gaithersburg only has Unit, Bldg, Fl, Ste, Apt
                ' so not going to worry about those
                Select Case splitRWordArr(1)
                    Case "BSMT", "FRNT", "LBBY", "LOWR", "OFC", "PH", "REAR", "SIDE", "UPPR"
                        validatedAddress.Item(AddressKey.UnitType) = splitRWordArr(1)
                        validatedAddress.Item(AddressKey.unitWithNum) = splitRWordArr(1)
                        streetAddress = splitRWordArr(0)
                    Case Else
                        validatedAddress.Item(AddressKey.UnitNum) = splitRWordArr(1)
                        
                        Dim splitUnitArr() As String
                        splitUnitArr = RWordTrim(splitRWordArr(0))
                        
                        validatedAddress.Item(AddressKey.UnitType) = splitUnitArr(1)
                        validatedAddress.Item(AddressKey.unitWithNum) = validatedAddress.Item(AddressKey.UnitType) & _
                                                                        " " & validatedAddress.Item(AddressKey.UnitNum)
                        streetAddress = splitUnitArr(0)
                End Select
            End If
            
            validatedAddress.Item(AddressKey.streetAddress) = streetAddress
        Else
            ' sometimes uspsData object is not populated e.g. just "501 Frederick Ave"
            ' unable to find an example where city and state and zip are given but uspsData does not populate
            Dim streetNum As String
            Dim route As String ' prefix + street name + type + postfix

            Dim component As Variant
            For Each component In jsonResult.Item("result").Item("address").Item("addressComponents")
                If component.Item("componentType") = "street_number" Then
                    streetNum = component.Item("componentName").Item("text")
                ElseIf component.Item("componentType") = "route" Then
                     route = component.Item("componentName").Item("text")
                ElseIf component.Item("componentType") = "subpremise" Then
                    validatedAddress.Item(AddressKey.unitWithNum) = component.Item("componentName").Item("text")
                ElseIf component.Item("componentType") = "postal_code" Then
                    validatedAddress.Item(AddressKey.zip) = component.Item("componentName").Item("text")
                End If
            Next component

            validatedAddress.Item(AddressKey.streetAddress) = streetNum & " " & route
        End If

        validatedAddress.Item(AddressKey.minLongitude) = jsonResult.Item("result").Item("geocode") _
                                                         .Item("bounds").Item("low").Item("longitude")
        validatedAddress.Item(AddressKey.maxLongitude) = jsonResult.Item("result").Item("geocode") _
                                                         .Item("bounds").Item("high").Item("longitude")
        validatedAddress.Item(AddressKey.minLatitude) = jsonResult.Item("result").Item("geocode") _
                                                         .Item("bounds").Item("low").Item("latitude")
        validatedAddress.Item(AddressKey.maxLatitude) = jsonResult.Item("result").Item("geocode") _
                                                         .Item("bounds").Item("high").Item("latitude")
    End If

    Set googleValidateQuery = validatedAddress
    Exit Function

JsonError:
    MsgBox "Error parsing Google Validation response: " & JsonConverter.ConvertToJson(jsonResult)
    Set googleValidateQuery = validatedAddress
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
    AddrLookupURL = AddrLookupURL & record.GburgFormatRawAddress.Item(AddressKey.streetAddress)
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


