Attribute VB_Name = "TestKzWorksheet"
Option Explicit
Option Private Module

'Kz���W���[���ɏ����ꂽPublic��Sub��Function��Rubberduck���g���ă��j�b�g�e�X�g����

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




'@TestMethod("KzVerifyWorksheetExists(sheetName As String)���e�X�g����")
Private Sub Test_KzVerifyWorksheetExists()
    'Assert:
    Assert.IsTrue KzVerifyWorksheetExists("Sheet1")
    Assert.IsFalse KzVerifyWorksheetExists("No Such Worksheet")
End Sub


'@TestMethod("KzIfWorksheetExistsInWorkbook���e�X�g����itrue��Ԃ��ꍇ�j")
Private Sub Test_KzIfWorksheetExistsInWorkbook()
    'Assert:
    Assert.IsTrue KzIfWorksheetExistsInWorkbook(ThisWorkbook, "Sheet1")
    Assert.IsFalse KzIfWorksheetExistsInWorkbook(ThisWorkbook, "No Such Worksheet")
End Sub


'@TestMethod("KzDeleteWorksheetIfExists(sheetName As String)���e�X�g����")
Private Sub Test_KzDeleteWorkSheetIfExists()
    'Arrange
    ' �J�����g��Workbook�Ƀ��[�N�V�[�g��}������A
    ' �V�[�g�̖��O��Test_DeleteWorkSheetIfExists
    Dim wsName As String: wsName = "Test_KzDeleteWorksheetIfExists"
    Worksheets.Add(After:=Worksheets(Worksheets.Count)) _
        .Name = wsName
    'Act:
    ' �}���������[�N�V�[�g���폜����
    KzDeleteWorksheetIfExists (wsName)
    ' �ꎞ�I�ɑ}���������[�N�V�[�g�����͂⑶�݂��Ȃ����Ƃ��m�F����
    Assert.IsFalse KzVerifyWorksheetExists(wsName)
End Sub



'@TestMethod("KzImportWorksheetFromWorkbook�����j�b�g�e�X�g����")
Private Sub Test_KzImportWorksheetFromWorkbook()
    On Error GoTo TestFail
    'Arrange:
    Dim wbSource As Workbook: Set wbSource = ActiveWorkbook
    Dim sourceSheetName As String: sourceSheetName = "Sheet1"
    Dim targetSheetName As String: targetSheetName = "work"
    'Act
    Call KzImportWorksheetFromWorkbook(wbSource, sourceSheetName, ThisWorkbook, targetSheetName)
    'Assert
    
    'TearDown
    Call KzDeleteWorksheetIfExists(targetSheetName)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

