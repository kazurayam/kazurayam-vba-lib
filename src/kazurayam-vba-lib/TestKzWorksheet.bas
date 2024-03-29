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


'@TestMethod("KzIfWorksheetExistsInWorkbookをテストする（trueを返す場合）")
Private Sub Test_KzIfWorksheetExistsInWorkbook()
    'Assert:
    Assert.IsTrue KzIfWorksheetExistsInWorkbook(ThisWorkbook, "Sheet1")
    Assert.IsFalse KzIfWorksheetExistsInWorkbook(ThisWorkbook, "No Such Worksheet")
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



'@TestMethod("KzFetchWorksheetFromWorkbookをユニットテストする")
Private Sub Test_KzFetchWorksheetFromWorkbook()
    On Error GoTo TestFail
    'Arrange
    Dim sourceWorkbook As Workbook: Set sourceWorkbook = ThisWorkbook
    Dim sourceSheetName As String: sourceSheetName = "Sheet1"
    Dim targetWorkbook As Workbook: Set targetWorkbook = sourceWorkbook
    Dim targetSheetName As String: targetSheetName = "copy"
    'Act
    Call KzFetchWorksheetFromWorkbook(sourceWorkbook, sourceSheetName, targetWorkbook, targetSheetName)
    'Assert
    Assert.IsTrue KzVerifyWorksheetExists("copy")
    'TearDown
    Application.DisplayAlerts = False
    ActiveWorkbook.Worksheets(targetSheetName).Delete
    Application.DisplayAlerts = True
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub



'@TestMethod("KzFetchWorksheetFromWorkbookがErrを投げる場合")
Private Sub Test_KzFetchWorksheetFromWorkbook_shouldThrowErr()
    On Error GoTo TestFail
    'Arrange
    Dim sourceWorkbook As Workbook: Set sourceWorkbook = ThisWorkbook
    Dim sourceSheetName As String: sourceSheetName = "Sheet1"
    Dim targetWorkbook As Workbook: Set targetWorkbook = sourceWorkbook
    Dim targetSheetName As String: targetSheetName = "Sheet1"
    'Act
    Call KzFetchWorksheetFromWorkbook(sourceWorkbook, sourceSheetName, targetWorkbook, targetSheetName)
    'Assert
    Assert.IsTrue KzVerifyWorksheetExists("copy")
    'TearDown
    Application.DisplayAlerts = False
    ActiveWorkbook.Worksheets(targetSheetName).Delete
    Application.DisplayAlerts = True
TestExit:
    Exit Sub
TestFail:
    'Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    'Errがraiseされるはずとわかっているので、何もしないで、静かに終了するべし
End Sub

