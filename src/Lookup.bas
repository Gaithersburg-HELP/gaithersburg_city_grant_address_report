Attribute VB_Name = "Lookup"
'@Folder("City_Grant_Address_Report.src")
Option Explicit

Public Enum addressKey
    streetAddress = -1 ' StreetNum + PrefixedStreetName + StreetType + Postfix
    Full = 0 ' StreetAddress + UnitType + UnitNum
    streetNum = 1
    PrefixedStreetname = 2 ' No postfix
    StreetType = 3
    Postfix = 4
    unitType = 5
    unitNum = 6
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
    address.Add addressKey.streetAddress, vbNullString
    address.Add addressKey.Full, vbNullString
    address.Add addressKey.streetNum, vbNullString
    address.Add addressKey.PrefixedStreetname, vbNullString
    address.Add addressKey.StreetType, vbNullString
    address.Add addressKey.Postfix, vbNullString
    address.Add addressKey.unitType, vbNullString
    address.Add addressKey.unitNum, vbNullString
    address.Add addressKey.unitWithNum, vbNullString
    address.Add addressKey.zip, vbNullString
    address.Add addressKey.minLongitude, 0
    address.Add addressKey.maxLongitude, 0
    address.Add addressKey.minLatitude, 0
    address.Add addressKey.maxLatitude, 0
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
            ' Google API quota limit is 429
            If .Status <> 429 Then Debug.Print "Error " + queryResult, vbCritical, "Connection"
            Set sendQuery = Nothing
        ElseIf .responseText <> vbNullString Then
            ' Debug.Print .responseText
            Set sendQuery = JsonConverter.ParseJson(.responseText)
        Else
            Set sendQuery = Nothing
        End If
    End With
End Function

' Executes REST query on Gaithersburg ArcGIS website Core_Address
' Expects Gaithersburg address with no unit type or number
' Returns number of results
Public Function gburgPartialQuery(ByVal fullAddress As String) As Long
    ' ' is escaped as ''
    Dim formatAddress As String
    formatAddress = Replace(fullAddress, "'", "''")

    
    Dim queryString As String
    queryString = "https://maps.gaithersburgmd.gov/arcgis/rest/services/layers/GaithersburgCityAddresses/MapServer/0/query?" & _
        "f=json&" & "returnGeometry=false&" & _
        "outFields=Full_Address,Address_Number,Road_Prefix_Dir,Road_Name,Road_Type,Road_Post_Dir,Unit_Type,Unit_Number,Zip_Code&" & _
        "where=Core_Address%20LIKE%20%27" & _
        WorksheetFunction.EncodeURL(formatAddress) & "%27"
           
    Dim jsonResult As Scripting.Dictionary
    Set jsonResult = sendQuery("GET", queryString, "application/x-www-form-urlencoded", vbNullString)
    
    If jsonResult Is Nothing Then
        gburgPartialQuery = 0
    Else
        gburgPartialQuery = jsonResult.Item("features").count
    End If
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
    
    If (Not (jsonResult Is Nothing)) And jsonResult.Item("features").count > 0 Then
        ' Since searching on Full_Address, expect only one feature to be returned
        Dim gburgAddress As Scripting.Dictionary
        Set gburgAddress = jsonResult.Item("features").Item(1).Item("attributes")
        
        validatedAddress.Item(addressKey.streetNum) = gburgAddress.Item("Address_Number")
        
        If gburgAddress.Item("Road_Prefix_Dir") <> vbNullString Then
            validatedAddress.Item(addressKey.PrefixedStreetname) = gburgAddress.Item("Road_Prefix_Dir") & " " & _
                                                                   gburgAddress.Item("Road_Name")
        Else
            validatedAddress.Item(addressKey.PrefixedStreetname) = gburgAddress.Item("Road_Name")
        End If
        
        validatedAddress.Item(addressKey.StreetType) = gburgAddress.Item("Road_Type")
        validatedAddress.Item(addressKey.Postfix) = gburgAddress.Item("Road_Post_Dir")
        validatedAddress.Item(addressKey.unitType) = gburgAddress.Item("Unit_Type")
        validatedAddress.Item(addressKey.unitNum) = gburgAddress.Item("Unit_Number")
        validatedAddress.Item(addressKey.zip) = gburgAddress.Item("Zip_Code")
        
        validatedAddress.Item(addressKey.streetAddress) = validatedAddress.Item(addressKey.streetNum) & " " & _
                                                          validatedAddress.Item(addressKey.PrefixedStreetname) & " " & _
                                                          validatedAddress.Item(addressKey.StreetType)
        If validatedAddress.Item(addressKey.Postfix) <> vbNullString Then
            validatedAddress.Item(addressKey.streetAddress) = validatedAddress.Item(addressKey.streetAddress) & " " & _
                                                              validatedAddress.Item(addressKey.Postfix)
        End If
        
        If validatedAddress.Item(addressKey.unitType) <> vbNullString Then
            validatedAddress.Item(addressKey.unitWithNum) = validatedAddress.Item(addressKey.unitType) & " " & _
                                                            validatedAddress.Item(addressKey.unitNum)
            validatedAddress.Item(addressKey.Full) = validatedAddress.Item(addressKey.streetAddress) & " " & _
                                                     validatedAddress.Item(addressKey.unitWithNum)
        Else
            validatedAddress.Item(addressKey.Full) = validatedAddress.Item(addressKey.streetAddress)
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
' - Returns Nothing if unable to get a response
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

    If jsonResult Is Nothing Then
        Set googleValidateQuery = Nothing
        Exit Function
    End If
    
    On Error GoTo JsonError
    ' Go with USPS CASS address response first, cannot trust Google components
    ' E.g. 600 s frederik dr at 100, gburg, md, 20877, usa: Google components keeps "at 100" instead of Ste 100
    ' see https://issuetracker.google.com/issues/325302835

    If jsonResult.Item("result").exists("uspsData") Then
        Dim uspsFullAddress As String
        uspsFullAddress = jsonResult.Item("result").Item("uspsData").Item("standardizedAddress").Item("firstAddressLine")
        
        Dim validPrimary As Boolean
        validPrimary = False
        Dim validSecondary As Boolean
        validSecondary = False
        ' USPS will sometimes return secondary even if unable to verify
        Dim secondaryReturned As Boolean
        secondaryReturned = False
        
        ' For DPV confirmation values, https://developers.google.com/maps/documentation/address-validation/handle-us-address
        ' Short circuiting does not exist in VBA
        If jsonResult.Item("result").Item("uspsData").exists("dpvConfirmation") Then
            Select Case jsonResult.Item("result").Item("uspsData").Item("dpvConfirmation")
                Case "Y"
                    validatedAddress.Item(addressKey.Full) = uspsFullAddress
                    validPrimary = True
                    
                    If jsonResult.Item("result").Item("verdict").Item("validationGranularity") = "SUB_PREMISE" Then
                        validSecondary = True
                        secondaryReturned = True
                    End If
                Case "S"
                    validPrimary = True
                    secondaryReturned = True
                Case "D"
                    validPrimary = True
            End Select
        End If
        
        ' Cannot rely on hasInferredComponents to check if address changed (always true when ZIP extension is not provided)
        ' Cannot rely on hasReplacedComponents either, only true if e.g. city is replaced
        ' google "replaced" or "spellCorrected" is not always present
        '   i.e. "fredrick" doesn't show spellcorrected but "frederik" does
        
        validatedAddress.Item(addressKey.zip) = jsonResult.Item("result").Item("uspsData") _
                                                          .Item("standardizedAddress").Item("zipCode")
        
        ' TODO? If not dpv confirmed, Get list of street names from Gaithersburg, Autocorrect to closest street name
        
        Dim streetAddress As String
        streetAddress = vbNullString
        
        If validPrimary Then streetAddress = uspsFullAddress
        
        If secondaryReturned Then
            ' see https://public-dhhs.ne.gov/nfocus/HowDoI/howdoi/usps_address_unit_types.htm
            
            Dim splitRWordArr() As String
            splitRWordArr = RWordTrim(uspsFullAddress)
            
            ' several designators such as upper don't have a secondary number
            ' however, as of 2/8/24 Gaithersburg only has Unit, Bldg, Fl, Ste, Apt
            ' so not going to worry about those
            Select Case splitRWordArr(1)
                Case "BSMT", "FRNT", "LBBY", "LOWR", "OFC", "PH", "REAR", "SIDE", "UPPR"
                    If validSecondary Then
                        validatedAddress.Item(addressKey.unitType) = splitRWordArr(1)
                        validatedAddress.Item(addressKey.unitWithNum) = splitRWordArr(1)
                    End If
                    streetAddress = splitRWordArr(0)
                Case Else
                    Dim splitUnitArr() As String
                    splitUnitArr = RWordTrim(splitRWordArr(0))

                    If validSecondary Then
                        
                        ' As of 2/27/24, USPS returns 150 Chevy Chase St Apt 102 # Unt when given Unt 102
                        Dim actualUnit As String
                        Dim actualNum As String
                        
                        Select Case splitUnitArr(1)
                            Case "APT", "BLDG", "DEPT", "FL", "HNGR", "KEY", "LOT", _
                                 "PIER", "SLIP", "SPC", "STE", "STOP", "TRLR", "UNIT"
                                actualUnit = splitUnitArr(1)
                                actualNum = splitRWordArr(1)
                                streetAddress = splitUnitArr(0)
                            Case Else ' "BSMT", "FRNT", "LBBY", "LOWR", "OFC", "PH", "REAR", "RM", "SIDE", "UPPR" "FRNT", "LOWR",
                                Dim secondSplitArr() As String
                                secondSplitArr = RWordTrim(splitUnitArr(0))
                                ' Gaithersburg database doesn't have Rm 1, etc. so drop last two words
                                ' actualNum = secondSplitArr(1) & " " & splitUnitArr(1) & " " & splitRWordArr(1)
                                actualNum = secondSplitArr(1)
                                Dim thirdSplitArr() As String
                                thirdSplitArr = RWordTrim(secondSplitArr(0))
                                actualUnit = thirdSplitArr(1)
                                streetAddress = thirdSplitArr(0)
                        End Select
                        validatedAddress.Item(addressKey.unitType) = actualUnit
                        validatedAddress.Item(addressKey.unitNum) = actualNum
                        validatedAddress.Item(addressKey.unitWithNum) = validatedAddress.Item(addressKey.unitType) & _
                                                                        " " & validatedAddress.Item(addressKey.unitNum)
                    Else
                        streetAddress = splitUnitArr(0)
                    End If
            End Select
        End If
        
        validatedAddress.Item(addressKey.streetAddress) = streetAddress
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
                validatedAddress.Item(addressKey.unitWithNum) = component.Item("componentName").Item("text")
            ElseIf component.Item("componentType") = "postal_code" Then
                validatedAddress.Item(addressKey.zip) = component.Item("componentName").Item("text")
            End If
        Next component

        validatedAddress.Item(addressKey.streetAddress) = streetNum & " " & route
    End If

    validatedAddress.Item(addressKey.minLongitude) = jsonResult.Item("result").Item("geocode") _
                                                     .Item("bounds").Item("low").Item("longitude")
    validatedAddress.Item(addressKey.maxLongitude) = jsonResult.Item("result").Item("geocode") _
                                                     .Item("bounds").Item("high").Item("longitude")
    validatedAddress.Item(addressKey.minLatitude) = jsonResult.Item("result").Item("geocode") _
                                                     .Item("bounds").Item("low").Item("latitude")
    validatedAddress.Item(addressKey.maxLatitude) = jsonResult.Item("result").Item("geocode") _
                                                     .Item("bounds").Item("high").Item("latitude")
    
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
    
    possibleInGburgQuery = (Not (jsonResult Is Nothing)) And jsonResult.Item("features").count > 0
End Function

