Attribute VB_Name = "FileUtilities"
Option Explicit

'@Folder "City_Grant_Address_Report.src"

' Expects a complete file path to CSV
' Returns a zero-based one dimensional array of each row in CSV
Public Function getCSV(ByVal path As String) As String()
    On Error GoTo CSVError
    
    Dim fileSO As FileSystemObject
    Dim fileTS As TextStream
    Set fileSO = New FileSystemObject
    Set fileTS = fileSO.OpenTextFile(path, ForReading, False, TristateUseDefault)
    
    Dim fileArr() As String
    fileArr = Split(fileTS.ReadAll, vbNewLine)
    getCSV = fileArr
    Exit Function
    
CSVError:
    Err.Raise 513
End Function

' Expects a CSV formatted as "places_api_old,apikey"
' Returns a dictionary of API keys
Public Function getAPIKeys() As Dictionary
    On Error GoTo APIerror
    
    Dim apiKeysDict As Object
    Set apiKeysDict = CreateObject("Scripting.Dictionary")
        
    Dim apiFileArray() As String
    apiFileArray = getCSV(ThisWorkbook.path & "\apikeys.csv")
    
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
    Err.Raise 513, Description:="invalid apikeys.csv, cannot continue"
End Function

