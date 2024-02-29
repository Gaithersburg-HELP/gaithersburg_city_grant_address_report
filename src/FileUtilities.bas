Attribute VB_Name = "FileUtilities"
Option Explicit

Public Const addressValidationKeyname As String = "address_key"

'@Folder "City_Grant_Address_Report.src"

Public Function getWorkbook() As Workbook
    Dim fd As FileDialog
    Dim selectedFile As String

    ' Create a FileDialog object
    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    ' Customize the FileDialog (optional)
    With fd
        .Title = "Select a File"  ' Dialog box title
        .AllowMultiSelect = False  ' Allow only single file selection
        .InitialFileName = ThisWorkbook.path   ' Set a default starting folder
        ' Add file filters if needed (Example)
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsm"
    End With

    ' Show the dialog. If user selects a file:
    If fd.Show = -1 Then
        selectedFile = fd.SelectedItems(1)  ' Get the selected file's path
        ' Do something with the file, e.g., open it:
        Set getWorkbook = Workbooks.Open(Filename:=selectedFile)
    Else
        Set getWorkbook = Nothing
    End If
End Function

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

' Expects a CSV formatted as "keyname,apikey"
' Returns a dictionary of API keys
Public Function getAPIKeys() As Scripting.Dictionary
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
    
    If Not apiKeysDict.Exists(addressValidationKeyname) Then
        Err.Raise 513
    End If
    
    Set getAPIKeys = apiKeysDict
    Exit Function

APIerror:
    Err.Raise 513, Description:="invalid apikeys.csv, cannot continue"
End Function

