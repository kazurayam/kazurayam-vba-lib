Attribute VB_Name = "TestKzWorksheet"
Option Explicit
Option Private Module

'Kzモジュールに書かれたPublicなSubやFunctionをRubberduckを使ってユニットテストする

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







'@TestMethod("KzVerifyWorksheetExists(sheetName As String)をテストする")
Private Sub Test_KzVerifyWorksheetExists()
    'Assert:
    Assert.IsTrue KzVerifyWorksheetExists("Sheet1")
    Assert.IsFalse KzVerifyWorksheetExists("No Such Worksheet")
End Sub

'@TestMethod("KzDeleteWorksheetIfExists(sheetName As String)をテストする")
Private Sub Test_KzDeleteWorkSheetIfExists()
    'Arrange
    ' カレントのWorkbookにワークシートを挿入する、
    ' シートの名前はTest_DeleteWorkSheetIfExists
    Dim wsName As String: wsName = "Test_KzDeleteWorksheetIfExists"
    Worksheets.Add(After:=Worksheets(Worksheets.Count)) _
        .Name = wsName
    'Act:
    ' 挿入したワークシートを削除する
    KzDeleteWorksheetIfExists (wsName)
    ' 一時的に挿入したワークシートがもはや存在しないことを確認する
    Assert.IsFalse KzVerifyWorksheetExists(wsName)
End Sub

'@TestMethod("KzExistsKey関数をユニットテストする")
Private Sub Test_KzExistsKey()
    'VBAのCollectionは連想配列のようにKeyとItemのペアを持つ場合もある
    '連想配列のようなCollectionが指定のKeyを持っているかどうかを調べてBooleanを返す
    'Arrange:
    Dim oCol As New Collection
    With oCol
        .Add Key:="テレビ", Item:="TV"
        .Add Key:="冷蔵庫", Item:="fridge"
        .Add Key:="炊飯器", Item:="rice cooker"
    End With
    'Assert
    Assert.IsTrue KzExistsKey(oCol, "炊飯器")
    Assert.IsFalse KzExistsKey(oCol, "ルンバ")
End Sub


