Attribute VB_Name = "RecordTupleUnitTest"
'@TestModule
'@Folder "City_Grant_Address_Report.test"

Option Explicit
Option Private Module

Private Assert As Object

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
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
    
    Assert.IsTrue record.visitData.Exists("food")
    Assert.IsTrue record.visitData.Item("food").Exists("Q1")
    Assert.IsTrue record.visitData.Item("food").Exists("Q2")
    Assert.IsTrue record.visitData.Item("food").Item("Q1")(1) = CDate("09/10/2023")
    Assert.IsTrue record.visitData.Item("food").Item("Q1")(2) = CDate("08/17/2023")
    Assert.IsTrue record.visitData.Item("food").Item("Q2")(1) = CDate("10/20/2024")
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
    
    Assert.IsTrue record.visitData.Exists("food")
    Assert.IsTrue record.visitData.Item("food").Exists("Q1")
    Assert.IsTrue record.visitData.Item("food").Exists("Q3")
    Assert.IsTrue record.visitData.Item("food").Exists("Q4")
    Assert.IsTrue record.visitData.Item("food").Item("Q1").Item(1) = "8/31/2023"
    Assert.IsTrue record.visitData.Item("food").Item("Q1").Item(2) = "9/15/2023"
End Sub

'@TestMethod
Public Sub TestFormatAddress()
    Dim record As RecordTuple
    Set record = New RecordTuple
    
    record.RawAddress = "501A S Frederick Ave E"
    record.RawUnitWithNum = "Suite 1"
    
    Assert.IsTrue record.isCorrectableAddress()
    
    Dim gburgFormat As Scripting.Dictionary
    Set gburgFormat = record.GburgFormatRawAddress
    
    Assert.IsTrue gburgFormat.Item(addressKey.Full) = "501a S Frederick Ave E Ste 1", "Full address incorrect"
    Assert.IsTrue gburgFormat.Item(addressKey.Postfix) = "E", "Postfix incorrect"
    Assert.IsTrue gburgFormat.Item(addressKey.PrefixedStreetName) = "S Frederick", "Street name incorrect"
    Assert.IsTrue gburgFormat.Item(addressKey.streetNum) = "501a", "Street no. incorrect"
    Assert.IsTrue gburgFormat.Item(addressKey.StreetType) = "Ave", "Street type incorrect"
    Assert.IsTrue gburgFormat.Item(addressKey.unitNum) = "1", "Unit no. incorrect"
    Assert.IsTrue gburgFormat.Item(addressKey.UnitType) = "Ste", "Unit type incorrect"
    
    Dim recordNoPostfix As RecordTuple
    Set recordNoPostfix = New RecordTuple
    
    recordNoPostfix.RawAddress = "2 Nina Ave"
    Set gburgFormat = recordNoPostfix.GburgFormatRawAddress
    
    Assert.IsTrue gburgFormat.Item(addressKey.Postfix) = vbNullString, "Postfix incorrect"
    Assert.IsTrue gburgFormat.Item(addressKey.PrefixedStreetName) = "Nina", "Street name incorrect"
    Assert.IsTrue gburgFormat.Item(addressKey.unitNum) = vbNullString, "Unit no. incorrect"
    Assert.IsTrue gburgFormat.Item(addressKey.UnitType) = vbNullString, "Unit type incorrect"
    
    Dim numericRecord As RecordTuple
    Set numericRecord = New RecordTuple
    
    numericRecord.RawAddress = "3458"
    Assert.IsFalse numericRecord.isCorrectableAddress(), "Numeric record marked as correctable"
    
    Dim alphabeticRecord As RecordTuple
    Set alphabeticRecord = New RecordTuple
    
    alphabeticRecord.RawAddress = "Asdfcvn Dfdwer"
    Assert.IsFalse alphabeticRecord.isCorrectableAddress(), "Alphabetic record marked as correctable"
End Sub

'@TestMethod
Public Sub TestIsAutocorrected()
    Dim record As RecordTuple
    Set record = New RecordTuple
    record.RawZip = "20878"
    record.RawAddress = "123 Test"
    record.RawUnitWithNum = "Apt 23"
    Assert.IsFalse record.isAutocorrected
    
    record.ValidZipcode = "20878"
    record.validAddress = "123 Test"
    record.validUnitWithNum = "Apt 23"
    Assert.IsFalse record.isAutocorrected
    
    record.validUnitWithNum = "Ste 23"
    Assert.IsTrue record.isAutocorrected

    record.validUnitWithNum = "Apt 23"
    record.ValidZipcode = "20877"
    Assert.IsTrue record.isAutocorrected

    record.ValidZipcode = "20878"
    record.validAddress = "124 test"
    Assert.IsTrue record.isAutocorrected
End Sub
