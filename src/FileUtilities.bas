Attribute VB_Name = "FileUtilities"
Option Explicit

Public Const addressValidationKeyname As String = "address_key"

'@Folder "City_Grant_Address_Report.src"

Public Function getWorkbook() As Workbook
    Dim fDialog As FileDialog
    Dim selectedFile As String

    ' Create a FileDialog object
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)

    ' Customize the FileDialog (optional)
    With fDialog
        .Title = "Select a File"  ' Dialog box title
        .allowMultiSelect = False  ' Allow only single file selection
        .InitialFileName = LibFileTools.GetLocalPath(ThisWorkbook.path)   ' Set a default starting folder
        ' Add file filters if needed (Example)
        .filters.Clear
        .filters.Add "Excel Files", "*.xlsm"
    End With

    ' Show the dialog. If user selects a file:
    If fDialog.Show = -1 Then
        selectedFile = fDialog.SelectedItems.Item(1)  ' Get the selected file's path
        ' Do something with the file, e.g., open it:
        Set getWorkbook = Workbooks.Open(fileName:=selectedFile)
    Else
        Set getWorkbook = Nothing
    End If
End Function

'@EntryPoint
Public Sub sortWorkbooks()
    Dim fDialog As FileDialog
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    With fDialog
        .allowMultiSelect = True
        .Title = "Select CSVs to Sort"
        .filters.Clear
        .filters.Add "CSV", "*.csv"
        .Show
    End With
    
    Dim selectedFile As Variant
    
    For Each selectedFile In fDialog.SelectedItems
        Dim wbook As Workbook
        Set wbook = Workbooks.Open(selectedFile)
        wbook.Activate
        
        Dim rng As Range
        Set rng = wbook.Worksheets.[_Default](1).UsedRange
        If rng.rows.count > 1 Then
            Set rng = rng.Resize(rng.rows.count - 1).Offset(1, 0)
            SheetUtilities.SortRange rng, True ' False
        End If
        wbook.Save
        wbook.Close
    Next selectedFile
End Sub

' Expects a complete file path to CSV
' Returns a zero-based one dimensional array of each row in CSV
Public Function getCSV(ByVal path As String) As String()
    On Error GoTo CSVError
    
    Dim fileSO As FileSystemObject
    Dim fileTS As TextStream
    Set fileSO = New FileSystemObject
    Set fileTS = fileSO.OpenTextFile(LibFileTools.GetLocalPath(path), ForReading, False, TristateUseDefault)
    
    Dim fileArr() As String
    fileArr = Split(fileTS.ReadAll, vbNewLine)
    getCSV = fileArr
    Exit Function
    
CSVError:
    Err.Raise 513
End Function

' Returns a dictionary of API keys
Public Function getAPIKeys() As Scripting.Dictionary
    On Error GoTo APIerror
    
    Dim apiKeysDict As Object
    Set apiKeysDict = CreateObject("Scripting.Dictionary")
    Dim addressKey As String
    addressKey = getAPIKeyRng().value
    
    If addressKey = vbNullString Then
        Err.Raise 513
    End If
    
    apiKeysDict.Add addressValidationKeyname, addressKey
    
    Set getAPIKeys = apiKeysDict
    Exit Function

APIerror:
    Err.Raise 513, Description:="no API address key, cannot continue"
End Function

