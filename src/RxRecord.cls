VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RxRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule EncapsulatePublicField,ProcedureNotUsed
'ProcedureNotUsed is Rubberduck bug
'@Folder("City_Grant_Address_Report.src")
Option Explicit

Private pQuarters(1 To 4) As Boolean

'Rubberduck bug
'@Ignore VariableNotUsed
Public guestID As String

Public medications As Scripting.Dictionary

Private Sub Class_Initialize()
    pQuarters(1) = False
    pQuarters(2) = False
    pQuarters(3) = False
    pQuarters(4) = False
    guestID = vbNullString
    Set medications = New Scripting.Dictionary
End Sub

Public Property Let quarter(ByVal quart As QuarterNum, ByVal present As Boolean)
    pQuarters(quart) = present
End Property

Public Property Get quarter(ByVal quart As QuarterNum) As Boolean
    quarter = pQuarters(quart)
End Property
