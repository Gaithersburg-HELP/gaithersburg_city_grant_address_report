VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RxTotals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("City_Grant_Address_Report.src")
Option Explicit

Public Enum QuarterNum
    '@Ignore UseMeaningfulName
    q1 = 1
    '@Ignore UseMeaningfulName
    q2 = 2
    '@Ignore UseMeaningfulName
    q3 = 3
    '@Ignore UseMeaningfulName
    q4 = 4
End Enum

Private Enum RxTotalRowNum
    numUndupResidentsServedRxRow = 1
    numUndupRxInProgramServedRow = 2
    numUndupRxInGburgServedRow = 3
    costRxInProgramRow = 4
    costRxInGburgRow = 5
    mostRecentRxDateRow = 6
    discardedIDsRow = 7
    [_RxTotalRowFirst] = numUndupResidentsServedRxRow
    [_RxTotalRowLast] = discardedIDsRow
End Enum

Private pTotals As Scripting.Dictionary
' Totals structure:
    ' { RxTotalType : { 1 (Quarter) : 0 (Total), 2 (Quarter) : 0 (Total), ... } }

Private Property Get totals() As Scripting.Dictionary
    If pTotals Is Nothing Then
        Set pTotals = New Scripting.Dictionary
        
        Dim rxTotalTypeRow As RxTotalRowNum
        For rxTotalTypeRow = [_RxTotalRowFirst] To [_RxTotalRowLast]
            If (rxTotalTypeRow = discardedIDsRow) Or (rxTotalTypeRow = mostRecentRxDateRow) Then
                pTotals.Add rxTotalTypeRow, vbNullString
            Else
                Dim Quarters As Scripting.Dictionary
                Set Quarters = New Scripting.Dictionary
                Quarters.Item(q1) = 0
                Quarters.Item(q2) = 0
                Quarters.Item(q3) = 0
                Quarters.Item(q4) = 0
                pTotals.Add rxTotalTypeRow, SheetUtilities.cloneDict(Quarters)
            End If
        Next rxTotalTypeRow
    End If
    Set totals = pTotals
End Property

Public Property Let numUndupResidentsServedRx(ByVal quarter As QuarterNum, ByVal total As Long)
    totals.Item(RxTotalRowNum.numUndupResidentsServedRxRow).Item(quarter) = total
End Property

Public Property Get numUndupResidentsServedRx(ByVal quarter As QuarterNum) As Long
    numUndupResidentsServedRx = totals.Item(RxTotalRowNum.numUndupResidentsServedRxRow).Item(quarter)
End Property

Public Property Let numUndupRxInProgramServed(ByVal quarter As QuarterNum, ByVal total As Long)
    totals.Item(numUndupRxInProgramServedRow).Item(quarter) = total
End Property

Public Property Get numUndupRxInProgramServed(ByVal quarter As QuarterNum) As Long
    numUndupRxInProgramServed = totals.Item(numUndupRxInProgramServedRow).Item(quarter)
End Property

Public Property Let numUndupRxInGburgServed(ByVal quarter As QuarterNum, ByVal total As Long)
    totals.Item(numUndupRxInGburgServedRow).Item(quarter) = total
End Property

Public Property Get numUndupRxInGburgServed(ByVal quarter As QuarterNum) As Long
    numUndupRxInGburgServed = totals.Item(numUndupRxInGburgServedRow).Item(quarter)
End Property

Public Property Let costRxInProgram(ByVal quarter As QuarterNum, ByVal total As Double)
    totals.Item(costRxInProgramRow).Item(quarter) = total
End Property

Public Property Get costRxInProgram(ByVal quarter As QuarterNum) As Double
    costRxInProgram = totals.Item(costRxInProgramRow).Item(quarter)
End Property

Public Property Let costRxInGburg(ByVal quarter As QuarterNum, ByVal total As Double)
    totals.Item(costRxInGburgRow).Item(quarter) = total
End Property

Public Property Get costRxInGburg(ByVal quarter As QuarterNum) As Double
    costRxInGburg = totals.Item(costRxInGburgRow).Item(quarter)
End Property

Public Property Let mostRecentRxDate(ByVal dateStr As String)
    totals.Item(mostRecentRxDateRow) = dateStr
End Property

Public Property Get mostRecentRxDate() As String
    mostRecentRxDate = totals.Item(mostRecentRxDateRow)
End Property

Public Property Let discardedIDs(ByVal ids As String)
    totals.Item(discardedIDsRow) = ids
End Property

Public Property Get discardedIDs() As String
    discardedIDs = totals.Item(discardedIDsRow)
End Property

Public Sub output()
    Dim rng As Range
    Set rng = SheetUtilities.getRxTotalsRng
    
    Dim rxTotalTypeRow As RxTotalRowNum
    For rxTotalTypeRow = [_RxTotalRowFirst] To [_RxTotalRowLast]
        If (rxTotalTypeRow = mostRecentRxDateRow) Then
            SheetUtilities.getRxMostRecentDateRng.value = totals.Item(rxTotalTypeRow)
        ElseIf (rxTotalTypeRow = discardedIDsRow) Then
            SheetUtilities.getRxDiscardedIDsRng.value = totals.Item(rxTotalTypeRow)
        Else
            Dim i As Long
            For i = 1 To 4
                rng.Cells.Item(rxTotalTypeRow, i) = totals.Item(rxTotalTypeRow).Item(i)
            Next i
        End If
    Next rxTotalTypeRow
End Sub
