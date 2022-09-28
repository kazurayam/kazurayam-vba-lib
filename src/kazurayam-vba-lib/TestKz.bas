Attribute VB_Name = "TestKz"
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

'@TestMethod("AbsolutifyPath�֐����e�X�g����")
Private Sub Test_KzAbsolutifyPath()
    '�t�@�C���̑��΃p�X���΃p�X�ɕϊ�����
    'Arrange:
    Dim base As String: base = ThisWorkbook.path
    Const relativePath = "..\Book1.xlsx"
    'Act:
    Dim absPath As String: absPath = KzAbsolutifyPath(base, relativePath)
    'Assert:
    Debug.Print "base: " & base
    Debug.Print "relative: " & relativePath
    Debug.Print "absPath: " & absPath
    Assert.IsTrue absPath Like "C:\*"         ' ��΃p�X�Ȃ�C:\�Ŏn�܂���
    Assert.IsTrue absPath Like "*\Book1.xlsx" ' \Book1.xlsx�ŏI���͂�
End Sub


'@TestMethod("KzToLocalFilePath()�����j�b�g�e�X�g����")
Private Sub Test_KzToLocalFilePath()
    'ToLocalFilePath��OneDrive�Ƀ}�b�s���O�����https://�Ŏn�܂�URL�ɑΉ�����t�@�C����C:\�Ŏn�܂郍�[�J���t�@�C���̃p�X�̕�����ɕϊ�����
    'Arrange:
    Dim Source As String: Source = "https://d.docs.live.net/c5960fe753e170b9/�f�X�N�g�b�v/Excel-Word-VBA"
    Dim expect As String: expect = "C:\Users\uraya\OneDrive\�f�X�N�g�b�v\Excel-Word-VBA"
    Dim actual As String
    'Act:
    actual = KzToLocalFilePath(Source)
    'Assert
    Debug.Print "source:" & vbTab; Chr(34); Source; Chr(34)
    Debug.Print "expect:" & vbTab; Chr(34); expect; Chr(34)
    Debug.Print "actual:" & vbTab; Chr(34); actual; Chr(34);
    Assert.IsTrue Len(actual) > 0
    Assert.IsTrue StrComp(expect, actual) = 0
End Sub


'@TestMethod("KzCreateFolder�֐������j�b�g�e�X�g����")
Private Sub Test_KzCreateFolder()
    'KzCreateFolder(p)�͈����Ƃ��Ďw�肳�ꂽp���t�H���_�̃p�X�Ƃ݂Ȃ��Ă��̃t�H���_�����B
    'p���܂�������ΐV�������Bp�����łɂ�������Ȃɂ����Ȃ��B
    'p�̐e�t�H���_���܂�������΃G���[�B�e�t�H���_�����ɂ�EnsureFolder�֐����g���B
    Dim p As String: p = ThisWorkbook.path & "\" & "tmp"
    KzCreateFolder (p)
    Assert.IsTrue KzPathExists(p)
    KzDeleteFolder (p)
End Sub

'@TestMethod("KzEnsureFolders�֐������j�b�g�e�X�g����")
Private Sub Test_KzEnsureFolders()
    'EnsureFolders(p)�̓t�H���_�����Bp�̐e�t�H���_�����������炻�̑c��ɂ܂ők���č��B
    Dim p As String: p = ThisWorkbook.path & "\build\tmp\testOutput"
    KzEnsureFolders (p)
    Assert.IsTrue KzPathExists(p)
    KzDeleteFolder (p)
End Sub

'@TestMethod("KzPathExists�֐������j�b�g�e�X�g����")
Private Sub Test_KzPathExists()
    Assert.IsTrue KzPathExists(ThisWorkbook.path)
    Dim p As String: p = ThisWorkbook.path & "\" & "kazurayam-vba-lib.xlsm"
    Assert.IsTrue KzPathExists(p)
End Sub

'@TestMethod("KzWriteTextIntoFile�֐���DeleteFile�֐����e�X�g����")
Private Sub Test_KzWriteTextIntoFile_and_KzDeleteFile()
    'Arrange:
    Dim folder As String: folder = ThisWorkbook.path & "\build"
    Dim file As String: file = folder & "\hello.txt"
    'Act:
    Call KzWriteTextIntoFile("Hello, world", file)
    'Assert:
    Debug.Assert KzPathExists(file)
    'TearDown
    KzDeleteFile (file)
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


