Attribute VB_Name = "Records"
Option Explicit

'@Folder "City_Grant_Address_Report.src"
'@EntryPoint
Public Sub confirmAddRecords()
    Dim confirmResponse As VbMsgBoxResult
    confirmResponse = MsgBox("Are you sure you wish to add records?", vbYesNo + vbQuestion, "Confirmation")
    If confirmResponse = vbNo Then
        Exit Sub
    End If
    
    addRecords
End Sub
' Returns autocorrected address, address valid json, autocorrect json
Public Function autocorrectAddress(ByVal address As String) As String()
    ' TODO write test for this function
    ' TODO Submit street name + Gaithersburg city only to place autocomplete
    ' ? Get list of street names from Gaithersburg, Autocorrect to closest street name
    ' Autocorrect Av to Ave, W Deer Pk to W Deer Park Rd
    ' Check postfixes
    autocorrectAddress = Array(address, "valid json", "autocorrect json")
End Function
Private Function getQuarter(ByVal dateStr As String) As String
    Select Case Month(dateStr)
        Case 7 To 9
            getQuarter = "Q1"
        Case 10 To 12
            getQuarter = "Q2"
        Case 1 To 3
            getQuarter = "Q3"
        Case 4 To 6
            getQuarter = "Q4"
    End Select
End Function

Private Sub writeToAddresses(ByVal record As RecordTuple)
    ' TODO write to addresses
    'getQuarter
    Dim testDict As Scripting.Dictionary
    Set testDict = New Scripting.Dictionary
    testDict.Add "Key1", "Basic value"
    
    Dim testArr() As String
    testArr = Split("arr1,arr2", ",")
    testDict.Add "Key2", testArr
    
    Debug.Print ConvertToJson(testDict)
End Sub

'@Ignore ParameterCanBeByVal
Private Sub tryAutocorrectRecord(ByRef addressDict As Dictionary, ByRef discardDict As Dictionary, ByRef record As RecordTuple)
    ' TODO autocorrecting
    ' autocorrectAddress(address)
    ' If autocorrected address is valid
        ' run against gaithersburg db
        ' Write to autocorrected addresses with json, highlight diff in yellow
        ' Add to address dictionary with gaithersburg result
    ' Else
        'add to discards dict, write to discards with autocorrect json
        ' If street name is in Gaithersburg street names
            ' highlight red
End Sub

' Prints Collection, checks if Collection contains JSON
'@Ignore ParameterCanBeByVal
Public Sub PrintCollection(ByRef collectionResult As Collection)
    Dim i As Long
    For i = 1 To collectionResult.Count
        If TypeOf collectionResult.Item(i) Is Dictionary Then
            PrintJson collectionResult.Item(i)
        ElseIf TypeOf collectionResult.Item(i) Is Collection Then
            PrintCollection collectionResult.Item(i)
        Else
            Debug.Print (collectionResult.Item(i) & ",");
        End If
    Next

End Sub

' Prints JSON
'@Ignore ParameterCanBeByVal
Public Sub PrintJson(ByRef jsonResult As Dictionary)
    Dim key As Variant
    For Each key In jsonResult
        If TypeOf jsonResult.Item(key) Is Collection Then
            Debug.Print (key & ": ")
            PrintCollection jsonResult.Item(key)
        ElseIf TypeOf jsonResult.Item(key) Is Dictionary Then
            PrintJson jsonResult.Item(key)
        Else
            Debug.Print (key & ": " & CStr(jsonResult.Item(key)))
        End If
    Next
End Sub

' Executes REST query on Gaithersburg ArcGIS website to see if address is in city or not
' Address should be given unencoded, in Proper Case
' Returns the number of results
' - 0 would be no results, 1 would be exact match, >2 would be multiple matches
Public Function ExecuteQuery(ByVal field As String, ByVal address As String) As Long
    ' ' is escaped as ''
    Dim formatAddress As String
    formatAddress = Replace(address, "'", "''")

    Dim service As Object
    Set service = New MSXML2.XMLHTTP60
    Dim queryString As String
    queryString = "https://maps.gaithersburgmd.gov/arcgis/rest/services/layers/GaithersburgCityAddresses/MapServer/0/query?" & _
        "f=json&" & "returnGeometry=false&" & _
        "outFields=OBJECTID,Address_Number,Road_Name,Road_Type,Full_Address&" & _
        "where=" & field & "%20LIKE%20%27" & _
        WorksheetFunction.EncodeURL(formatAddress) & "%27"
    
    With service
        .Open "GET", queryString, False
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        .send
        Do While .readyState <> 4
            DoEvents
        Loop
        
        Dim queryResult As String
        If .Status >= 400 And .Status <= 599 Then
            queryResult = CStr(.Status) + " - " + .statusText
            If .responseText <> vbNullString Then
                queryResult = queryResult + vbCrLf & .responseText
            End If
            MsgBox "Error " + queryResult, vbCritical, "Connection"
            ExecuteQuery = 0
            Exit Function
        End If
        
        queryResult = .responseText
    End With

    Dim jsonResult As Scripting.Dictionary
    
    If queryResult <> vbNullString Then
        Set jsonResult = JsonConverter.ParseJson(queryResult)
        If Not jsonResult Is Nothing Then
            ExecuteQuery = jsonResult.Item("features").Count
        Else
            ExecuteQuery = 0
        End If
    Else
        ExecuteQuery = 0
    End If
End Function


Private Function loadRecord(ByVal recordRow As Range) As RecordTuple
    Dim record As RecordTuple
    Set record = New RecordTuple
    
    record.VisitDate = recordRow.Value
    record.service = recordRow.Offset(0, 1).Value
    record.GuestID = recordRow.Offset(0, 2).Value
    record.FirstName = recordRow.Offset(0, 3).Value
    record.LastName = recordRow.Offset(0, 4).Value
    record.RawAddress = recordRow.Offset(0, 5).Value
    record.Apt = recordRow.Offset(0, 6).Value
    record.City = recordRow.Offset(0, 7).Value
    record.State = recordRow.Offset(0, 8).Value
    record.Zip = recordRow.Offset(0, 9).Value
    record.HouseholdTotal = recordRow.Offset(0, 10).Value
    record.RxTotal = recordRow.Offset(0, 11).Value
    
    Set loadRecord = record
End Function
Public Sub addRecords()
    ' TODO import REST, json helper functions and json converter from Module 1
    ' TODO import MicroTimer from Module 1
    Dim currentRecord As RecordTuple
    
    ' Load addresses into dictionary
    ' Initialize discards dictionary
    
    ' Save application status bar, update status bar with current record progress
    Dim recordProgress As Long
    recordProgress = 1
    Dim appStatus As Variant
    If Application.StatusBar = False Then appStatus = False Else appStatus = Application.StatusBar
    
    ' LOOP through records until hit last row with data
    ' ws.Cells(ws.Rows.Count, column).End(xlUp).Row
        ' loadRecord(row)
        ' If row is blank or is not correctable, remove row, next row
        ' If address is in discards dictionary
            'Write to discards
        ' Else If address is in address dictionary
            ' Merge service and visit data to address dictionary
        ' Else
            ' Run against gaithersburg database
            ' If in gaithersburg database, add to address dictionary with Gaithersburg autocorrect
            ' curRecord.getGburgFormatRawAddress ("FullAddr")
            Select Case ExecuteQuery("Core_Address", vbNullString)
                Case 0
                    'InCity = ""
                Case 1
                    'InCity = "Yes"
                Case Else
                    ' If multiple matches, probably an apartment building
                    'InCity = "Yes"
            End Select
        
        ' update status bar, yield execution so Excel remains responsive and user can hit Esc
        Application.StatusBar = "Processed record " & recordProgress
        recordProgress = recordProgress + 1
        DoEvents
    
    ' Clear all addresses
    ' writeToAddresses
    ' Compute totals
    
    Application.StatusBar = appStatus
End Sub


