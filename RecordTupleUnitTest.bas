Attribute VB_Name = "RecordTupleUnitTest"
'@TestModule
'@Folder "City_Grant_Address_Report.test"

Option Explicit
Option Private Module

Private assert As Object

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set assert = CreateObject("Rubberduck.AssertClass")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set assert = Nothing
End Sub

'@TestMethod
Public Sub TestMergeRecord()
    Dim record As RecordTuple
    Set record = New RecordTuple
    Dim recordToMerge As RecordTuple
    Set recordToMerge = New RecordTuple
    
    record.AddVisit "food", "09/10/2023"
    record.AddVisit "food", "08/17/2023"
    recordToMerge.AddVisit "food", "10/20/2024"
    
    record.MergeRecordVisitData recordToMerge
    
    assert.IsTrue record.visitData.Exists("food")
    assert.IsTrue record.visitData.Item("food").Exists("Q1")
    assert.IsTrue record.visitData.Item("food").Exists("Q2")
    assert.IsTrue record.visitData.Item("food").Item("Q1")(1) = "09/10/2023"
    assert.IsTrue record.visitData.Item("food").Item("Q1")(2) = "08/17/2023"
    assert.IsTrue record.visitData.Item("food").Item("Q2")(1) = "10/20/2024"
End Sub

'@TestMethod
Public Sub TestVisitJson()
    Dim record As RecordTuple
    Set record = New RecordTuple
    
    Dim visitData As Scripting.Dictionary
    Set visitData = New Scripting.Dictionary
    
    visitData.Add "food", JsonConverter.ParseJson( _
        "{""Q1"":[""8/31/2023"",""9/15/2023""],""Q3"":[""2/15/2023""],""Q4"":[""5/31/2023""]}")
    
    Set record.visitData = visitData
    
    assert.IsTrue record.visitData.Exists("food")
    assert.IsTrue record.visitData.Item("food").Exists("Q1")
    assert.IsTrue record.visitData.Item("food").Exists("Q3")
    assert.IsTrue record.visitData.Item("food").Exists("Q4")
    assert.IsTrue record.visitData.Item("food").Item("Q1").Item(1) = "8/31/2023"
    assert.IsTrue record.visitData.Item("food").Item("Q1").Item(2) = "9/15/2023"
End Sub

'@TestMethod
Public Sub TestFormatAddress()
    Dim record As RecordTuple
    Set record = New RecordTuple
    
    record.RawAddress = "501A S Frederick Ave E"
    record.RawUnitWithNum = "Suite 1"
    
    assert.IsTrue record.isCorrectableAddress()
    
    Dim gburgFormat As Scripting.Dictionary
    Set gburgFormat = record.getGburgFormatRawAddress()
    
    assert.IsTrue gburgFormat.Item(record.FullAddressKey) = "501a S Frederick Ave E", "Full address incorrect"
    assert.IsTrue gburgFormat.Item(record.PostfixKey) = "E", "Postfix incorrect"
    assert.IsTrue gburgFormat.Item(record.PrefixedStreetNameKey) = "S Frederick", "Street name incorrect"
    assert.IsTrue gburgFormat.Item(record.StreetNumKey) = "501a", "Street no. incorrect"
    assert.IsTrue gburgFormat.Item(record.StreetTypeKey) = "Ave", "Street type incorrect"
    assert.IsTrue gburgFormat.Item(record.UnitNumKey) = "1", "Unit no. incorrect"
    assert.IsTrue gburgFormat.Item(record.UnitTypeKey) = "Ste", "Unit type incorrect"
    
    Dim recordNoPostfix As RecordTuple
    Set recordNoPostfix = New RecordTuple
    
    recordNoPostfix.RawAddress = "2 Nina Ave"
    Set gburgFormat = recordNoPostfix.getGburgFormatRawAddress()
    
    assert.IsTrue gburgFormat.Item(record.PostfixKey) = vbNullString, "Postfix incorrect"
    assert.IsTrue gburgFormat.Item(record.PrefixedStreetNameKey) = "Nina", "Street name incorrect"
    assert.IsTrue gburgFormat.Item(record.UnitNumKey) = vbNullString, "Unit no. incorrect"
    assert.IsTrue gburgFormat.Item(record.UnitTypeKey) = vbNullString, "Unit type incorrect"
    
    Dim numericRecord As RecordTuple
    Set numericRecord = New RecordTuple
    
    numericRecord.RawAddress = "3458"
    assert.IsFalse numericRecord.isCorrectableAddress(), "Numeric record marked as correctable"
    
    Dim alphabeticRecord As RecordTuple
    Set alphabeticRecord = New RecordTuple
    
    alphabeticRecord.RawAddress = "Asdfcvn Dfdwer"
    assert.IsFalse alphabeticRecord.isCorrectableAddress(), "Alphabetic record marked as correctable"
End Sub
