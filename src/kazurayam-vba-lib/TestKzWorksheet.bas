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

'@TestMethod("KzExistsKey�֐������j�b�g�e�X�g����")
Private Sub Test_KzExistsKey()
    'VBA��Collection�͘A�z�z��̂悤��Key��Item�̃y�A�����ꍇ������
    '�A�z�z��̂悤��Collection���w���Key�������Ă��邩�ǂ����𒲂ׂ�Boolean��Ԃ�
    'Arrange:
    Dim oCol As New Collection
    With oCol
        .Add Key:="�e���r", Item:="TV"
        .Add Key:="�①��", Item:="fridge"
        .Add Key:="���ъ�", Item:="rice cooker"
    End With
    'Assert
    Assert.IsTrue KzExistsKey(oCol, "���ъ�")
    Assert.IsFalse KzExistsKey(oCol, "�����o")
End Sub


