Attribute VB_Name = "FileUtilities"
Option Explicit

'@Folder "City_Grant_Address_Report.src"
' Expects a CSV formatted as "places_api_old,apikey"
' Returns a dictionary of API keys
Public Function getAPIKeys() As Variant
    On Error GoTo APIerror
    
    Dim apiKeysDict As Object
    Set apiKeysDict = CreateObject("Scripting.Dictionary")
        
    Dim apiFileSO As FileSystemObject
    Dim apiFileTS As TextStream
    Set apiFileSO = New FileSystemObject
    Set apiFileTS = apiFileSO.OpenTextFile(ThisWorkbook.Path & "\apikeys.csv", ForReading, False, TristateUseDefault)
    
    Dim apiFileArray() As String
    apiFileArray = Split(apiFileTS.ReadAll, vbNewLine)
    
    Dim i As Long
    Dim apiFileArrLine() As String
    For i = LBound(apiFileArray, 1) To UBound(apiFileArray, 1)
        apiFileArrLine = Split(apiFileArray(i), ",")
        apiKeysDict.Add apiFileArrLine(0), apiFileArrLine(1)
    Next i
    
    If Not apiKeysDict.Exists("places_api_old") Then
        Err.Raise 513
    End If
    
    Set getAPIKeys = apiKeysDict
    Exit Function

APIerror:
    MsgBox ("invalid apikeys.csv, cannot continue")
    Set getAPIKeys = Nothing
    Exit Function
    
End Function

