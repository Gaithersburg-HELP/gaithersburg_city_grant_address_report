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
    
    record.AddVisit "09/10/2023", "food"
    record.AddVisit "08/17/2023", "food"
    recordToMerge.AddVisit "10/20/2024", "food"
    
    record.MergeRecord recordToMerge
    
    assert.IsTrue record.visitData.Exists("food")
    assert.IsTrue record.visitData.Item("food").Exists("Q1")
    assert.IsTrue record.visitData.Item("food").Exists("Q2")
    assert.IsTrue record.visitData.Item("food").Item("Q1")(1) = CDate("09/10/2023")
    assert.IsTrue record.visitData.Item("food").Item("Q1")(2) = CDate("08/17/2023")
    assert.IsTrue record.visitData.Item("food").Item("Q2")(1) = CDate("10/20/2024")
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
    Set gburgFormat = record.GburgFormatRawAddress
    
    assert.IsTrue gburgFormat.Item(AddressKey.Full) = "501a S Frederick Ave E Ste 1", "Full address incorrect"
    assert.IsTrue gburgFormat.Item(AddressKey.Postfix) = "E", "Postfix incorrect"
    assert.IsTrue gburgFormat.Item(AddressKey.PrefixedStreetName) = "S Frederick", "Street name incorrect"
    assert.IsTrue gburgFormat.Item(AddressKey.streetNum) = "501a", "Street no. incorrect"
    assert.IsTrue gburgFormat.Item(AddressKey.StreetType) = "Ave", "Street type incorrect"
    assert.IsTrue gburgFormat.Item(AddressKey.unitNum) = "1", "Unit no. incorrect"
    assert.IsTrue gburgFormat.Item(AddressKey.UnitType) = "Ste", "Unit type incorrect"
    
    Dim recordNoPostfix As RecordTuple
    Set recordNoPostfix = New RecordTuple
    
    recordNoPostfix.RawAddress = "2 Nina Ave"
    Set gburgFormat = recordNoPostfix.GburgFormatRawAddress
    
    assert.IsTrue gburgFormat.Item(AddressKey.Postfix) = vbNullString, "Postfix incorrect"
    assert.IsTrue gburgFormat.Item(AddressKey.PrefixedStreetName) = "Nina", "Street name incorrect"
    assert.IsTrue gburgFormat.Item(AddressKey.unitNum) = vbNullString, "Unit no. incorrect"
    assert.IsTrue gburgFormat.Item(AddressKey.UnitType) = vbNullString, "Unit type incorrect"
    
    Dim numericRecord As RecordTuple
    Set numericRecord = New RecordTuple
    
    numericRecord.RawAddress = "3458"
    assert.IsFalse numericRecord.isCorrectableAddress(), "Numeric record marked as correctable"
    
    Dim alphabeticRecord As RecordTuple
    Set alphabeticRecord = New RecordTuple
    
    alphabeticRecord.RawAddress = "Asdfcvn Dfdwer"
    assert.IsFalse alphabeticRecord.isCorrectableAddress(), "Alphabetic record marked as correctable"
End Sub

'@TestMethod
Public Sub TestIsAutocorrected()
    Dim record As RecordTuple
    Set record = New RecordTuple
    record.RawZip = "20878"
    record.RawAddress = "123 Test"
    record.RawUnitWithNum = "Apt 23"
    assert.IsFalse record.isAutocorrected
    
    record.ValidZipcode = "20878"
    record.ValidAddress = "123 Test"
    record.validUnitWithNum = "Apt 23"
    assert.IsFalse record.isAutocorrected
    
    record.validUnitWithNum = "Ste 23"
    assert.IsTrue record.isAutocorrected

    record.validUnitWithNum = "Apt 23"
    record.ValidZipcode = "20877"
    assert.IsTrue record.isAutocorrected

    record.ValidZipcode = "20878"
    record.ValidAddress = "124 test"
    assert.IsTrue record.isAutocorrected
End Sub
