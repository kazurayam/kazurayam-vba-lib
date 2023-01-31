Attribute VB_Name = "TestKzRange"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("KzGetUniqueItems�֐����e�X�g����")
Private Sub Test_KzGetUniqueItems()
    Debug.Print String(300, vbCrLf)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("Sheet1")
    Dim tbl As ListObject: Set tbl = ws.ListObjects("�e�[�u��1")
    Dim item As Variant
    Dim uniqueItems As Variant
    uniqueItems = KzRange.KzGetUniqueItems(tbl.ListColumns(1).DataBodyRange)
    Dim i As Long: i = 0
    For Each item In uniqueItems
        i = i + 1
        Debug.Print i & " " & item
    Next

    Assert.AreEqual CLng(2), i
End Sub
