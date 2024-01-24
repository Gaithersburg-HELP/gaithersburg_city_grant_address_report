VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RecordTuple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("City_Grant_Address_Report.src")
Option Explicit

Private pVisitDate As String
Private pService As String
Private pGuestID As String
Private pFirstName As String
Private pLastName As String
Private pRawAddress As String
Private pApt As String
Private pCity As String
Private pState As String
Private pZip As String
Private pHouseholdTotal As Long
Private pRxTotal As Double

Private pValidAddress As String
Private pValidZipcode As String

'@Ignore MoveFieldCloserToUsage
Private gburgFormatRawAddress As Scripting.Dictionary
'@Ignore MoveFieldCloserToUsage
Private gburgFormatValidAddress As Scripting.Dictionary

Public Property Get VisitDate() As String
    VisitDate = pVisitDate
End Property

Public Property Let VisitDate(ByVal str As String)
    pVisitDate = str
End Property

Public Property Get service() As String
    service = pService
End Property

Public Property Let service(ByVal str As String)
    pService = str
End Property

Public Property Get GuestID() As String
    GuestID = pGuestID
End Property

Public Property Let GuestID(ByVal str As String)
    pGuestID = str
End Property

Public Property Get FirstName() As String
    FirstName = pFirstName
End Property

Public Property Let FirstName(ByVal str As String)
    pFirstName = str
End Property

Public Property Get LastName() As String
    LastName = pLastName
End Property

Public Property Let LastName(ByVal str As String)
    pLastName = str
End Property

Public Property Get RawAddress() As String
    RawAddress = pRawAddress
End Property

Public Property Let RawAddress(ByVal str As String)
    pRawAddress = str
End Property

Public Property Get Apt() As String
    Apt = pApt
End Property

Public Property Let Apt(ByVal str As String)
    pApt = str
End Property

Public Property Get City() As String
    City = pCity
End Property

Public Property Let City(ByVal str As String)
    pCity = str
End Property

Public Property Get State() As String
    State = pState
End Property

Public Property Let State(ByVal str As String)
    pState = str
End Property

Public Property Get Zip() As String
    Zip = pZip
End Property

Public Property Let Zip(ByVal str As String)
    pZip = str
End Property

Public Property Get HouseholdTotal() As Long
    HouseholdTotal = pHouseholdTotal
End Property

Public Property Let HouseholdTotal(ByVal num As Long)
    pHouseholdTotal = num
End Property

Public Property Get RxTotal() As Double
    RxTotal = pRxTotal
End Property

Public Property Let RxTotal(ByVal num As Double)
    pRxTotal = num
End Property

Public Property Get ValidAddress() As String
    ValidAddress = pValidAddress
End Property

Public Property Let ValidAddress(ByVal str As String)
    pValidAddress = str
End Property

Public Property Get ValidZipcode() As String
    ValidZipcode = pValidZipcode
End Property

Public Property Let ValidZipcode(ByVal str As String)
    pValidZipcode = str
End Property

Public Property Get CleanFirstName() As String
    CleanFirstName = StrConv(FirstName, vbProperCase)
End Property

Public Property Get CleanLastName() As String
    Dim properLastName As String
    properLastName = LastName
    properLastName = StrConv(LastName, vbProperCase)
    ' Last names beginning with "Mc" should have next letter capitalized
    If Left$(properLastName, 2) = "Mc" Then
        properLastName = "Mc" & UCase$(Mid$(properLastName, 3, 1)) & Mid$(properLastName, 4, 99)
    End If
    CleanLastName = properLastName
End Property

Public Property Get CleanFullName() As String
    CleanFullName = CleanFirstName & " " & CleanLastName
End Property

Public Property Get CleanInitials() As String
    CleanInitials = UCase$(Left$(CleanFirstName, 1)) _
                    & UCase$(Left$(CleanLastName, 1))
End Property

' Trims off first word including space afterwards
' Returns [trimmed first word, trimmed string (blank if only one word)]
Private Function LWordTrim(ByVal str As String) As String()
    Dim firstWord As String
    Dim spaceIndex As Long
    spaceIndex = InStr(1, str, " ", vbTextCompare)
    If (spaceIndex <> 0) Then
        firstWord = Left$(str, spaceIndex - 1)
        LWordTrim = Array(firstWord, Right$(str, Len(str) - spaceIndex))
    Else
        LWordTrim = Array(str, vbNullString)
    End If
End Function

' Trims off last word including space before
' Returns [trimmed string, trimmed last word (blank if only one word)]
Private Function RWordTrim(ByVal str As String) As String()
    Dim lastWord As String
    Dim spaceIndex As Long
    spaceIndex = InStrRev(str, " ", -1, vbTextCompare)
    If (spaceIndex <> 0) Then
        lastWord = Right$(str, Len(str) - spaceIndex)
        RWordTrim = Array(Left$(str, spaceIndex - 1), lastWord)
    Else
        RWordTrim = Array(str, vbNullString)
    End If
End Function

' Checks street name from start position onwards for an abbreviated direction with spaces around it
' If direction is found, returns it, otherwise returns ""
'@Ignore AssignedByValParameter
Private Function GetDirection(ByVal start As Long, ByVal street As String) As String
    If (street = "  ") Then
        GetDirection = vbNullString
        Exit Function
    End If
    If (start < 1) Then
        start = 1
    End If
    
    ' Add spaces to avoid e.g. Westland
    street = " " & street & " "
    If InStr(start, street, " NE ", vbTextCompare) Then
        GetDirection = "NE"
    ElseIf InStr(start, street, " NW ", vbTextCompare) Then
        GetDirection = "NW"
    ElseIf InStr(start, street, " SW ", vbTextCompare) Then
        GetDirection = "SW"
    ElseIf InStr(start, street, " SE ", vbTextCompare) Then
        GetDirection = "SE"
    ElseIf InStr(start, street, " N ", vbTextCompare) Then
        GetDirection = "N"
    ElseIf InStr(start, street, " E ", vbTextCompare) Then
        GetDirection = "E"
    ElseIf InStr(start, street, " S ", vbTextCompare) Then
        GetDirection = "S"
    ElseIf InStr(start, street, " W ", vbTextCompare) Then
        GetDirection = "W"
    Else
        GetDirection = vbNullString
    End If
End Function


Public Property Get StreetNum() As String
    StreetNum = "StreetNum"
End Property
Public Property Get PrefixedStreetName() As String
    PrefixedStreetName = "PrefixedStreetName"
End Property
Public Property Get StreetType() As String
    StreetType = "StreetType"
End Property
Public Property Get Postfix() As String
    Postfix = "Postfix"
End Property
Public Property Get FullAddress() As String
    FullAddress = "FullAddress"
End Property

' Returns dictionary of StreetNum, PrefixedStreetName, StreetType, Postfix, FullAddress
' - capitalized Street Name with abbr. prefix, cap. abbr. Street Type, cap. abbr. Postfix
'@Ignore AssignedByValParameter
Private Function getFormattedSplitAddr(ByVal addr As String) As Scripting.Dictionary
    addr = Trim$(RawAddress)
    addr = Replace(addr, "   ", " ")
    addr = Replace(addr, "  ", " ")
    addr = Replace(addr, ".", vbNullString)
    
    addr = StrConv(addr, vbProperCase)
    
    ' - Gaithersburg has O'Neill and Odend'hal, but Soxbox exports as Oneill (and Oneil misspelling) and Odendhal
    ' - Mccausland Pl renamed to Prism Pl
    addr = Replace(addr, "Oneill", "O'neill")
    addr = Replace(addr, "Oneil", "O'neill")
    addr = Replace(addr, "Odendhal", "Odend'hal")
    addr = Replace(addr, "Mccausland", "Prism")
    
    ' Add spaces to left and right of street name (temporarily) for replace operations (avoid e.g. Westland Dr)
    addr = " " & addr & " "
    addr = Replace(addr, " Northwest ", " NW ")
    addr = Replace(addr, " Southwest ", " SW ")
    addr = Replace(addr, " Northeast ", " NE ")
    addr = Replace(addr, " Southeast ", " SE ")
    addr = Replace(addr, " North ", " N ")
    addr = Replace(addr, " South ", " S ")
    addr = Replace(addr, " East ", " E ")
    addr = Replace(addr, " West ", " W ")
    addr = Trim$(addr)
    
    addr = Replace(addr, "Avenue", "Ave")
    addr = Replace(addr, "Boulevard", "Blvd")
    addr = Replace(addr, "Circle", "Cir")
    addr = Replace(addr, "Close", "Cls")
    addr = Replace(addr, "Court", "Ct")
    addr = Replace(addr, "Drive", "Dr")
    addr = Replace(addr, "Highway", "Hwy")
    addr = Replace(addr, "Lane", "Ln")
    addr = Replace(addr, "Mews", "Mews")
    addr = Replace(addr, "Parkway", "Pkwy")
    addr = Replace(addr, "Place", "Pl")
    addr = Replace(addr, "Road", "Rd")
    addr = Replace(addr, "Square", "Sq")
    addr = Replace(addr, "Street", "St")
    addr = Replace(addr, "Terrace", "Ter")
    addr = Replace(addr, "Way", "Way")

    Dim splitAddr As Scripting.Dictionary
    Set splitAddr = New Scripting.Dictionary
    splitAddr.Add FullAddress, addr

    Dim splitNumArr() As String
    ' Split Street Number from Street name
    splitNumArr = LWordTrim(addr)
    splitAddr.Add StreetNum, splitNumArr(0)
    
    Dim streetPostfix As String
    streetPostfix = GetDirection(Len(addr) - 3, addr)
    Dim splitTypeArr() As String
    ' Remove postfix if it exists
    If streetPostfix <> vbNullString Then
        splitTypeArr = LWordTrim(LWordTrim(splitNumArr(1))(0))
    Else
        splitTypeArr = LWordTrim(splitNumArr(1))
    End If
    
    splitAddr.Add PrefixedStreetName, splitTypeArr(0)
    splitAddr.Add StreetType, splitTypeArr(1)
    splitAddr.Add Postfix, streetPostfix
    
    Set getFormattedSplitAddr = splitAddr
End Function

Public Function getGburgFormatRawAddress() As Scripting.Dictionary
    If gburgFormatRawAddress Is Nothing Then
        Set gburgFormatRawAddress = getFormattedSplitAddr(RawAddress)
    End If
    Set getGburgFormatRawAddress = gburgFormatRawAddress
End Function

Public Function getGburgFormatValidAddress() As Scripting.Dictionary
    If gburgFormatValidAddress Is Nothing Then
        Set gburgFormatValidAddress = getFormattedSplitAddr(ValidAddress)
    End If
    Set getGburgFormatValidAddress = gburgFormatValidAddress
End Function

' TODO check if is correctable
Public Function isCorrectableAddress() As Boolean
    isCorrectableAddress = True
End Function


