Attribute VB_Name = "TestG"
Option Explicit
Option Private Module

'G���W���[���ɏ����ꂽSub��Function��Rubberduck���g���ă��j�b�g�e�X�g����

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
Private Sub Test_AbsolutifyPath()
    '�t�@�C���̑��΃p�X���΃p�X�ɕϊ�����
    'Arrange:
    Dim base As String: base = ThisWorkbook.path
    Const relativePath = "..\Book1.xlsx"
    'Act:
    Dim absPath As String: absPath = AbsolutifyPath(base, relativePath)
    'Assert:
    Debug.Print "base: " & base
    Debug.Print "relative: " & relativePath
    Debug.Print "absPath: " & absPath
    Assert.IsTrue absPath Like "C:\*"         ' ��΃p�X�Ȃ�C:\�Ŏn�܂���
    Assert.IsTrue absPath Like "*\Book1.xlsx" ' \Book1.xlsx�ŏI���͂�
End Sub


'@TestMethod("ToLocalFilePath()�����j�b�g�e�X�g����")
Private Sub Test_ToLocalFilePath()
    'ToLocalFilePath��OneDrive�Ƀ}�b�s���O�����https://�Ŏn�܂�URL�ɑΉ�����t�@�C����C:\�Ŏn�܂郍�[�J���t�@�C���̃p�X�̕�����ɕϊ�����
    'Arrange:
    Dim Source As String: Source = "https://d.docs.live.net/c5960fe753e170b9/�f�X�N�g�b�v/Excel-Word-VBA"
    Dim expect As String: expect = "C:\Users\uraya\OneDrive\�f�X�N�g�b�v\Excel-Word-VBA"
    Dim actual As String
    'Act:
    actual = ToLocalFilePath(Source)
    'Assert
    Debug.Print "source:" & vbTab; Chr(34); Source; Chr(34)
    Debug.Print "expect:" & vbTab; Chr(34); expect; Chr(34)
    Debug.Print "actual:" & vbTab; Chr(34); actual; Chr(34);
    Assert.IsTrue Len(actual) > 0
    Assert.IsTrue StrComp(expect, actual) = 0
End Sub


'@TestMethod("CreateFolder�֐������j�b�g�e�X�g����")
Private Sub Test_CreateFolder()
    'CreateFolder(p)�͈����Ƃ��Ďw�肳�ꂽp���t�H���_�̃p�X�Ƃ݂Ȃ��Ă��̃t�H���_�����B
    'p���܂�������ΐV�������Bp�����łɂ�������Ȃɂ����Ȃ��B
    'p�̐e�t�H���_���܂�������΃G���[�B�e�t�H���_�����ɂ�EnsureFolder�֐����g���B
    Dim p As String: p = ThisWorkbook.path & "\" & "tmp"
    CreateFolder (p)
    Assert.IsTrue PathExists(p)
    DeleteFolder (p)
End Sub

'@TestMethod("EnsureFolders�֐������j�b�g�e�X�g����")
Private Sub Test_EnsureFolders()
    'EnsureFolders(p)�̓t�H���_�����Bp�̐e�t�H���_�����������炻�̑c��ɂ܂ők���č��B
    Dim p As String: p = ThisWorkbook.path & "\build\tmp\testOutput"
    EnsureFolders (p)
    Assert.IsTrue PathExists(p)
    DeleteFolder (p)
End Sub

'@TestMethod("PathExists�֐������j�b�g�e�X�g����")
Private Sub Test_PathExists()
    Assert.IsTrue PathExists(ThisWorkbook.path)
    Dim p As String: p = ThisWorkbook.path & "\" & "kazurayam-vba-lib.xlsm"
    Assert.IsTrue PathExists(p)
End Sub

'@TestMethod("WriteTextIntoFile�֐���DeleteFile�֐����e�X�g����")
Private Sub Test_WriteTextIntoFile_and_DeleteFile()
    'Arrange:
    Dim folder As String: folder = ThisWorkbook.path & "\build"
    Dim file As String: file = folder & "\hello.txt"
    'Act:
    Call WriteTextIntoFile("Hello, world", file)
    'Assert:
    Debug.Assert PathExists(file)
    'TearDown
    DeleteFile (file)
End Sub

'@TestMethod("VerifyWorksheetExists(sheetName As String)���e�X�g����")
Private Sub Test_VerifyWorksheetExists()
    'Assert:
    Assert.IsTrue VerifyWorksheetExists("Sheet1")
    Assert.IsFalse G.VerifyWorksheetExists("No Such Worksheet")
End Sub

'@TestMethod("DeleteWorksheetIfExists(sheetName As String)���e�X�g����")
Private Sub Test_DeleteWorkSheetIfExists()
    'Arrange
    ' �J�����g��Workbook�Ƀ��[�N�V�[�g��}������A
    ' �V�[�g�̖��O��Test_DeleteWorkSheetIfExists
    Dim wsName As String: wsName = "Test_DeleteWorksheetIfExists"
    Worksheets.Add(After:=Worksheets(Worksheets.Count)) _
        .Name = wsName
    'Act:
    ' �}���������[�N�V�[�g���폜����
    G.DeleteWorksheetIfExists (wsName)
    ' �ꎞ�I�ɑ}���������[�N�V�[�g�����͂⑶�݂��Ȃ����Ƃ��m�F����
    Assert.IsFalse G.VerifyWorksheetExists(wsName)
End Sub

'@TestMethod("ExistsKey�֐������j�b�g�e�X�g����")
Private Sub Test_ExistsKey()
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
    Assert.IsTrue ExistsKey(oCol, "���ъ�")
    Assert.IsFalse ExistsKey(oCol, "�����o")
End Sub


