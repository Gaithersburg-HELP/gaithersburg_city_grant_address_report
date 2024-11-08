VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RxRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule UseMeaningfulName, MoveFieldCloserToUsage, EncapsulatePublicField
'@Folder("City_Grant_Address_Report.src")
Option Explicit

Public quarter1 As Boolean
Public quarter2 As Boolean
Public quarter3 As Boolean
Public quarter4 As Boolean
Public guestID As String

Public medications As Scripting.Dictionary

Private Sub Class_Initialize()
    quarter1 = False
    quarter2 = False
    quarter3 = False
    quarter4 = False
    guestID = vbNullString
    Set medications = New Scripting.Dictionary
End Sub
