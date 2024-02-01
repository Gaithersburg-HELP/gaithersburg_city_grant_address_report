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

Private pInCity As String
Private pUserVerified As Boolean

Private pValidAddress As String
Private pValidUnitWithNum As String
Private pValidZipcode As String

Private pRawAddress As String
Private pRawUnitWithNum As String
Private pRawCity As String
Private pRawState As String
Private pRawZip As String
Private pGuestID As String
Private pFirstName As String
Private pLastName As String

'BUG does not track household total by quarter
Private pHouseholdTotal As Long

' Quarter: Dictionary (VisitDate Date: RxTotal Double)
Private pRxTotal As Scripting.Dictionary

' Service: Dictionary (Quarter: VisitDate Date Collection)
Private pVisitData As Scripting.Dictionary

Private pGburgFormatRawAddress As Scripting.Dictionary
Private pGburgFormatValidAddress As Scripting.Dictionary

Public Property Get key() As String
    key = pGuestID
End Property


Public Property Get InCity() As String
    InCity = pInCity
End Property

Public Property Let InCity(ByVal str As String)
    pInCity = str
End Property

Public Property Get UserVerified() As Boolean
    UserVerified = pUserVerified
End Property

Public Property Let UserVerified(ByVal bool As Boolean)
    pUserVerified = bool
End Property

Public Property Get ValidAddress() As String
    ValidAddress = pValidAddress
End Property

Public Property Let ValidAddress(ByVal str As String)
    pValidAddress = str
End Property

Public Property Get ValidUnitWithNum() As String
    ValidUnitWithNum = pValidUnitWithNum
End Property

Public Property Let ValidUnitWithNum(ByVal str As String)
    pValidUnitWithNum = str
End Property

Public Property Get FullValidAddress() As String
    If (ValidUnitWithNum <> vbNullString) Then
        FullValidAddress = ValidAddress & " " & ValidUnitWithNum
    Else
        FullValidAddress = ValidAddress
    End If
End Property

Public Property Get ValidZipcode() As String
    ValidZipcode = pValidZipcode
End Property

Public Property Let ValidZipcode(ByVal str As String)
    pValidZipcode = str
End Property

Public Property Get RawAddress() As String
    RawAddress = pRawAddress
End Property

Public Property Let RawAddress(ByVal str As String)
    If pRawAddress <> vbNullString Then
        Err.Raise Number:=513, Description:="Reassignment of raw address"
    End If
    
    pRawAddress = str
End Property

Public Property Get RawUnitWithNum() As String
    RawUnitWithNum = pRawUnitWithNum
End Property

Public Property Let RawUnitWithNum(ByVal str As String)
    pRawUnitWithNum = str
End Property

Public Property Get FullRawAddress() As String
    If (RawUnitWithNum <> vbNullString) Then
        FullRawAddress = RawAddress & " " & RawUnitWithNum
    Else
        FullRawAddress = RawAddress
    End If
End Property

Public Property Get RawCity() As String
    RawCity = pRawCity
End Property

Public Property Let RawCity(ByVal str As String)
    pRawCity = str
End Property

Public Property Get RawState() As String
    RawState = pRawState
End Property

Public Property Let RawState(ByVal str As String)
    pRawState = str
End Property

Public Property Get RawZip() As String
    RawZip = pRawZip
End Property

Public Property Let RawZip(ByVal str As String)
    pRawZip = str
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

Public Property Get householdTotal() As Long
    householdTotal = pHouseholdTotal
End Property

Public Property Let householdTotal(ByVal num As Long)
    pHouseholdTotal = num
End Property

Public Property Get rxTotal() As Scripting.Dictionary
    If pRxTotal Is Nothing Then Set pRxTotal = New Scripting.Dictionary
    Set rxTotal = pRxTotal
End Property

Public Property Set rxTotal(ByVal dict As Scripting.Dictionary)
    Set pRxTotal = dict
End Property

Public Property Get visitData() As Scripting.Dictionary
    If pVisitData Is Nothing Then Set pVisitData = New Scripting.Dictionary
    Set visitData = pVisitData
End Property

Public Property Set visitData(ByVal dict As Scripting.Dictionary)
    Set pVisitData = dict
End Property

Private Function getQuarterStr(ByVal dateStr As String) As String
    Select Case Month(dateStr)
        Case 7 To 9
            getQuarterStr = "Q1"
        Case 10 To 12
            getQuarterStr = "Q2"
        Case 1 To 3
            getQuarterStr = "Q3"
        Case 4 To 6
            getQuarterStr = "Q4"
    End Select
End Function

Public Sub AddVisit(ByVal visitDate As String, ByVal service As String)
    Dim newVisitData As Scripting.Dictionary
    Set newVisitData = New Scripting.Dictionary
    
    Dim serviceVisits As Scripting.Dictionary
    Set serviceVisits = New Scripting.Dictionary
    
    Dim visits As Collection
    Set visits = New Collection
    
    ' CStr formats Date more nicely for readability
    visits.Add CStr(CDate(visitDate))
    serviceVisits.Add getQuarterStr(visitDate), visits
    newVisitData.Add service, serviceVisits
    
    MergeVisitData newVisitData
End Sub

Public Sub addRx(ByVal visitDate As String, ByVal rx As Double)
    Dim newRx As Scripting.Dictionary
    Set newRx = New Scripting.Dictionary
    
    Dim rxVisits As Scripting.Dictionary
    Set rxVisits = New Scripting.Dictionary
    
    ' NOTE Adding Date as dictionary key silently changes string formatting on print for some reason
    rxVisits.Add CDate(visitDate), rx
    newRx.Add getQuarterStr(visitDate), rxVisits
    
    MergeRx newRx
End Sub

Public Sub MergeRecord(ByVal record As RecordTuple)
    MergeVisitData record.visitData
    MergeRx record.rxTotal
End Sub

Private Sub MergeRx(ByVal rxToMerge As Scripting.Dictionary)
    Dim quarter As Variant
    For Each quarter In rxToMerge.Keys
        If rxTotal.Exists(quarter) Then
            Dim visit As Variant
            For Each visit In rxToMerge.Item(quarter).Keys
                If Not rxTotal.Item(quarter).Exists(visit) Then
                    rxTotal.Item(quarter).Add visit, rxToMerge.Item(quarter).Item(visit)
                End If
            Next visit
        Else
            rxTotal.Add quarter, rxToMerge.Item(quarter)
        End If
    Next quarter
End Sub

Private Sub MergeVisitData(ByVal visitDataToMerge As Scripting.Dictionary)
    Dim service As Variant
    
    For Each service In visitDataToMerge.Keys
        If visitData.Exists(service) Then
            ' Loop through quarters and merge
            Dim quarter As Variant
            For Each quarter In visitDataToMerge.Item(service)
                If visitData.Item(service).Exists(quarter) Then
                    Dim visitDateToMerge As Variant
                    For Each visitDateToMerge In visitDataToMerge.Item(service).Item(quarter)
                        Dim visitExists As Boolean
                        Dim existingVisitDate As Variant
                        For Each existingVisitDate In visitData.Item(service).Item(quarter)
                            If visitDateToMerge = existingVisitDate Then
                                visitExists = True
                                Exit For
                            End If
                        Next existingVisitDate
                        
                        If Not visitExists Then
                            visitData.Item(service).Item(quarter).Add visitDateToMerge
                        End If
                    Next visitDateToMerge
                Else
                    visitData.Item(service).Add quarter, visitDataToMerge.Item(service).Item(quarter)
                End If
            Next quarter
        Else
            visitData.Add service, visitDataToMerge.Item(service)
        End If
    Next service
End Sub

Public Property Get CleanFirstName() As String
    CleanFirstName = StrConv(FirstName, vbProperCase)
End Property

Public Property Get CleanLastName() As String
    Dim properLastName As String
    properLastName = StrConv(LastName, vbProperCase)
    ' Last names beginning with "Mc" should have next letter capitalized
    If Left$(properLastName, 2) = "Mc" Then
        properLastName = "Mc" & UCase$(Mid$(properLastName, 3, 1)) & Mid$(properLastName, 4, 99)
    End If
    CleanLastName = properLastName
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
        LWordTrim = Split(firstWord & "|" & Right$(str, Len(str) - spaceIndex), "|")
    Else
        LWordTrim = Split(str & "|", "|")
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
        RWordTrim = Split(Left$(str, spaceIndex - 1) & "|" & lastWord, "|")
    Else
        RWordTrim = Split(str & "|", "|")
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

' Replaces 2 or 3 spaces with single space with nullstring, returns result as ProperCase
'@Ignore AssignedByValParameter
Private Function CleanString(ByVal str As String) As String
    str = Trim$(str)
    str = Replace(str, "   ", " ")
    str = Replace(str, "  ", " ")
    ' Street numbers can have periods, see https://pe.usps.com/text/pub28/28c2_013.htm
    CleanString = StrConv(str, vbProperCase)
End Function

' Splits unit into zero-based array of abbreviated unit type, unit num, additional words
' If only one word, assumes unit is Apt
'@Ignore AssignedByValParameter
Private Function SplitUnit(ByVal unitWithNum As String) As String()
    unitWithNum = CleanString(unitWithNum)
    
    ' See https://public-dhhs.ne.gov/nfocus/HowDoI/howdoi/usps_address_unit_types.htm
    unitWithNum = Replace(unitWithNum, "Apartment", "Apt")
    unitWithNum = Replace(unitWithNum, "Building", "Bldg")
    unitWithNum = Replace(unitWithNum, "Basement", "Bsmt")
    unitWithNum = Replace(unitWithNum, "Department", "Dept")
    unitWithNum = Replace(unitWithNum, "Floor", "Fl")
    unitWithNum = Replace(unitWithNum, "Front", "Frnt")
    unitWithNum = Replace(unitWithNum, "Hanger", "Hngr")
    unitWithNum = Replace(unitWithNum, "Key", "Key")
    unitWithNum = Replace(unitWithNum, "Lobby", "Lbby")
    unitWithNum = Replace(unitWithNum, "Lot", "Lot")
    unitWithNum = Replace(unitWithNum, "Lower", "Lowr")
    unitWithNum = Replace(unitWithNum, "Ofc", "Office")
    unitWithNum = Replace(unitWithNum, "Penthouse", "Ph")
    unitWithNum = Replace(unitWithNum, "Pier", "Pier")
    unitWithNum = Replace(unitWithNum, "Rear", "Rear")
    unitWithNum = Replace(unitWithNum, "Room", "Rm")
    unitWithNum = Replace(unitWithNum, "Side", "Side")
    unitWithNum = Replace(unitWithNum, "Slip", "Slip")
    unitWithNum = Replace(unitWithNum, "Space", "Spc")
    unitWithNum = Replace(unitWithNum, "Suite", "Ste")
    unitWithNum = Replace(unitWithNum, "Stop", "Stop")
    unitWithNum = Replace(unitWithNum, "Trailer", "Trlr")
    unitWithNum = Replace(unitWithNum, "Unit", "Unit")
    unitWithNum = Replace(unitWithNum, "Upper", "Uppr")
    
    ' If # assume Apt
    unitWithNum = Replace(unitWithNum, "#", "Apt")
    
    Dim splitArr() As String
    splitArr = Split(unitWithNum, " ")
    
    Select Case UBound(splitArr)
        Case 0
            ' If one word, assume Apt
            SplitUnit = Split("Apt," & unitWithNum, ",")
        Case -1
            ' unitWithNum was ""
            SplitUnit = Split(",", ",")
        Case Else
            SplitUnit = splitArr
    End Select
End Function

' Returns dictionary of all AddressKey enum
' - capitalized Street Name with abbr. prefix, cap. abbr. Street Type, cap. abbr. Postfix
'@Ignore AssignedByValParameter
Private Function getFormattedSplitAddr(ByVal addr As String) As Scripting.Dictionary
    addr = CleanString(addr)
    
    ' BUG these will not show up in autocorrected
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

    Dim splitNumArr() As String
    ' Split Street Number from Street name
    splitNumArr = LWordTrim(addr)
    splitAddr.Add AddressKey.StreetNum, splitNumArr(0)
    
    ' Remove periods, hyphens, not in street name
    ' see https://pe.usps.com/text/pub28/28c2_013.htm & https://pe.usps.com/text/dmm100/addressing-mail.htm
    splitNumArr(0) = Replace(splitNumArr(1), ".", vbNullString)
    splitNumArr(0) = Replace(splitNumArr(1), "-", vbNullString)
    
    Dim streetPostfix As String
    streetPostfix = GetDirection(Len(addr) - 3, addr)
    Dim splitTypeArr() As String
    ' Remove postfix if it exists
    If streetPostfix <> vbNullString Then
        splitTypeArr = RWordTrim(RWordTrim(splitNumArr(1))(0))
    Else
        splitTypeArr = RWordTrim(splitNumArr(1))
    End If
    
    splitAddr.Add AddressKey.PrefixedStreetName, splitTypeArr(0)
    splitAddr.Add AddressKey.StreetType, splitTypeArr(1)
    splitAddr.Add AddressKey.Postfix, streetPostfix
    
    Dim splitUnitArr() As String
    splitUnitArr = SplitUnit(RawUnitWithNum)
    
    splitAddr.Add AddressKey.UnitType, splitUnitArr(0)
    splitAddr.Add AddressKey.UnitNum, splitUnitArr(1)
    
    If splitUnitArr(0) <> vbNullString Then
        splitAddr.Add AddressKey.Full, addr & " " & splitUnitArr(0) & " " & splitUnitArr(1)
    Else
        splitAddr.Add AddressKey.Full, addr
    End If
    
    Set getFormattedSplitAddr = splitAddr
End Function

Public Property Get GburgFormatRawAddress() As Scripting.Dictionary
    If pGburgFormatRawAddress Is Nothing Then
        Set pGburgFormatRawAddress = getFormattedSplitAddr(RawAddress)
    End If
    Set GburgFormatRawAddress = pGburgFormatRawAddress
End Property

Public Property Get GburgFormatValidAddress() As Scripting.Dictionary
    If pGburgFormatValidAddress Is Nothing Then
        Set pGburgFormatValidAddress = getFormattedSplitAddr(ValidAddress)
    End If
    Set GburgFormatValidAddress = pGburgFormatValidAddress
End Property

Public Function isCorrectableAddress() As Boolean
    Dim splitAddr() As String
    splitAddr = Split(Trim$(RawAddress), " ")
    Select Case UBound(splitAddr)
        Case -1 'null
            isCorrectableAddress = False
        Case Else
            Dim hasNumber As Boolean
            Dim hasLetter As Boolean
            Dim i As Long
            For i = 1 To Len(RawAddress)
                Dim char As String
                char = Mid$(RawAddress, i, 1)
                If IsNumeric(char) Then hasNumber = True
                If Not IsNumeric(char) Then hasLetter = True
                If hasNumber And hasLetter Then Exit For
            Next
            isCorrectableAddress = hasNumber And hasLetter
    End Select
End Function


