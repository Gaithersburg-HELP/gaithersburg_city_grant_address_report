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

Public Function getQuarter(ByVal dateStr As String) As String
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
Private Function ExecuteQuery(ByVal field As String, ByVal address As String) As Long
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

Private Function loadRecordFromRaw(ByVal recordRowFirstCell As Range) As RecordTuple
    Dim record As RecordTuple
    Set record = New RecordTuple
    
    record.AddVisit recordRowFirstCell.Offset(0, 1).Value, recordRowFirstCell.Value
    record.UserVerified = False

    record.GuestID = recordRowFirstCell.Offset(0, 2).Value
    record.FirstName = recordRowFirstCell.Offset(0, 3).Value
    record.LastName = recordRowFirstCell.Offset(0, 4).Value
    record.RawAddress = recordRowFirstCell.Offset(0, 5).Value
    record.RawUnitWithNum = recordRowFirstCell.Offset(0, 6).Value
    record.RawCity = recordRowFirstCell.Offset(0, 7).Value
    record.RawState = recordRowFirstCell.Offset(0, 8).Value
    record.RawZip = recordRowFirstCell.Offset(0, 9).Value
    record.HouseholdTotal = recordRowFirstCell.Offset(0, 10).Value
    record.RxTotal = recordRowFirstCell.Offset(0, 11).Value
    
    Set loadRecordFromRaw = record
End Function

Private Function loadServiceNames(ByVal sheetName As String) As String()
    Dim servicesRng As Range
    Set servicesRng = SheetUtilities.getServiceHeaderRng(sheetName)
    ReDim services(servicesRng.Count) As String
    Dim i As Long
    i = 1
    Do While i <= servicesRng.Count
        services(i) = servicesRng.Cells.Item(1, i).Value
    Loop
    
    loadServiceNames = services
End Function

Private Function loadAddresses(ByVal sheetName As String) As Scripting.Dictionary
    Dim sheet As Worksheet
    Set sheet = ActiveWorkbook.Worksheets.[_Default](sheetName)
    
    Dim addresses As Scripting.Dictionary
    Set addresses = New Scripting.Dictionary
    
    If sheet.Range("A2").Value = vbNullString Then
        Set loadAddresses = addresses
        Exit Function
    End If
    
    Dim services() As String
    services = loadServiceNames(sheetName)
    
    Dim i As Long
    i = 1
    Do While i < getBlankRow(sheetName).row
        Dim recordRow As Range
        Set recordRow = sheet.Rows.Item(i).Cells
        
        Dim record As RecordTuple
        Set record = New RecordTuple
        
        record.InCity = recordRow.Cells.Item(1, 1).Value
        record.UserVerified = CBool(recordRow.Cells.Item(1, 2).Value)
        record.ValidAddress = recordRow.Cells.Item(1, 3).Value
        record.ValidUnitWithNum = recordRow.Cells.Item(1, 4).Value
        record.ValidZipcode = recordRow.Cells.Item(1, 5).Value
        record.RawAddress = recordRow.Cells.Item(1, 6).Value
        record.RawUnitWithNum = recordRow.Cells.Item(1, 7).Value
        record.RawCity = recordRow.Cells.Item(1, 8).Value
        record.RawState = recordRow.Cells.Item(1, 9).Value
        record.RawZip = recordRow.Cells.Item(1, 10).Value
        record.GuestID = recordRow.Cells.Item(1, 11).Value
        record.FirstName = recordRow.Cells.Item(1, 12).Value
        record.LastName = recordRow.Cells.Item(1, 13).Value
        record.HouseholdTotal = recordRow.Cells.Item(1, 14).Value
        record.RxTotal = recordRow.Cells.Item(1, 15).Value
        
        Dim visitData As Scripting.Dictionary
        Set visitData = New Scripting.Dictionary
        
        Dim j As Long
        j = 1
        Do While j <= UBound(services) + 1
            visitData.Add services(j - 1), JsonConverter.ParseJson(recordRow.Cells.Item(1, 15 + j).Value)
        Loop
        
        Set record.visitData = visitData
        
        addresses.Add record.FullRawAddress, record
        i = i + 1
    Loop

    Set loadAddresses = addresses
End Function

Private Sub writeAddress(ByVal sheetName As String, ByVal record As RecordTuple)
    Dim sheet As Worksheet
    Set sheet = ActiveWorkbook.Worksheets.[_Default](sheetName)
    
    Dim addresses As Scripting.Dictionary
    Set addresses = New Scripting.Dictionary
    
    Dim services() As String
    If sheet.Range("A2").Value <> vbNullString Then
        services = loadServiceNames(sheetName)
    End If

    Dim recordRow As Range
    Set recordRow = getBlankRow(sheetName)
    ' BUG pick up here
    recordRow.Cells(1, 1).Value = record.InCity
    
    Dim visitData As Scripting.Dictionary
    Set visitData = New Scripting.Dictionary
    
    Dim j As Long
    j = 1
    Do While j <= UBound(services) + 1
        visitData.Add services(j - 1), JsonConverter.ParseJson(recordRow.Cells.Item(1, 15 + j).Value)
    Loop
    
    Set record.visitData = visitData
    
    addresses.Add record.FullRawAddress, record
        
    Dim serviceColumn As Scripting.Dictionary
    ' If service doesn't yet exist, add service with column number
    
    ' Sort service columns
End Sub

Public Sub addRecords()
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
    ' getBlankRow
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


