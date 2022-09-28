Attribute VB_Name = "TestKzFile"
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

'@TestMethod("AbsolutifyPath関数をテストする")
Private Sub Test_KzAbsolutifyPath()
    'ファイルの相対パスを絶対パスに変換する
    'Arrange:
    Dim base As String: base = ThisWorkbook.path
    Const relativePath = "..\Book1.xlsx"
    'Act:
    Dim absPath As String: absPath = KzAbsolutifyPath(base, relativePath)
    'Assert:
    Debug.Print "base: " & base
    Debug.Print "relative: " & relativePath
    Debug.Print "absPath: " & absPath
    Assert.IsTrue absPath Like "C:\*"         ' 絶対パスならC:\で始まって
    Assert.IsTrue absPath Like "*\Book1.xlsx" ' \Book1.xlsxで終わるはず
End Sub


'@TestMethod("KzToLocalFilePath()をユニットテストする")
Private Sub Test_KzToLocalFilePath()
    'ToLocalFilePathはOneDriveにマッピングされてhttps://で始まるURLに対応するファイルをC:\で始まるローカルファイルのパスの文字列に変換する
    'Arrange:
    Dim Source As String: Source = "https://d.docs.live.net/c5960fe753e170b9/デスクトップ/Excel-Word-VBA"
    Dim expect As String: expect = "C:\Users\uraya\OneDrive\デスクトップ\Excel-Word-VBA"
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


'@TestMethod("KzCreateFolder関数をユニットテストする")
Private Sub Test_KzCreateFolder()
    'KzCreateFolder(p)は引数として指定されたpをフォルダのパスとみなしてそのフォルダを作る。
    'pがまだ無ければ新しく作る。pがすでにあったらなにもしない。
    'pの親フォルダがまだ無ければエラー。親フォルダを作るにはEnsureFolder関数を使え。
    Dim p As String: p = ThisWorkbook.path & "\" & "tmp"
    KzCreateFolder (p)
    Assert.IsTrue KzPathExists(p)
    KzDeleteFolder (p)
End Sub



'@TestMethod("KzEnsureFolders関数をユニットテストする")
Private Sub Test_KzEnsureFolders()
    'EnsureFolders(p)はフォルダを作る。pの親フォルダが無かったらその祖先にまで遡って作る。
    Dim p As String: p = ThisWorkbook.path & "\build\tmp\testOutput"
    KzEnsureFolders (p)
    Assert.IsTrue KzPathExists(p)
    KzDeleteFolder (p)
End Sub




'@TestMethod("KzPathExists関数をユニットテストする")
Private Sub Test_KzPathExists()
    Assert.IsTrue KzPathExists(ThisWorkbook.path)
    Dim p As String: p = ThisWorkbook.path & "\" & "kazurayam-vba-lib.xlsm"
    Assert.IsTrue KzPathExists(p)
End Sub


'@TestMethod("KzWriteTextIntoFile関数とDeleteFile関数をテストする")
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




