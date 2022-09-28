Attribute VB_Name = "TestG"
Option Explicit
Option Private Module

'Gモジュールに書かれたSubやFunctionをRubberduckを使ってユニットテストする

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
Private Sub Test_AbsolutifyPath()
    'ファイルの相対パスを絶対パスに変換する
    'Arrange:
    Dim base As String: base = ThisWorkbook.path
    Const relativePath = "..\Book1.xlsx"
    'Act:
    Dim absPath As String: absPath = AbsolutifyPath(base, relativePath)
    'Assert:
    Debug.Print "base: " & base
    Debug.Print "relative: " & relativePath
    Debug.Print "absPath: " & absPath
    Assert.IsTrue absPath Like "C:\*"         ' 絶対パスならC:\で始まって
    Assert.IsTrue absPath Like "*\Book1.xlsx" ' \Book1.xlsxで終わるはず
End Sub


'@TestMethod("ToLocalFilePath()をユニットテストする")
Private Sub Test_ToLocalFilePath()
    'ToLocalFilePathはOneDriveにマッピングされてhttps://で始まるURLに対応するファイルをC:\で始まるローカルファイルのパスの文字列に変換する
    'Arrange:
    Dim Source As String: Source = "https://d.docs.live.net/c5960fe753e170b9/デスクトップ/Excel-Word-VBA"
    Dim expect As String: expect = "C:\Users\uraya\OneDrive\デスクトップ\Excel-Word-VBA"
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


'@TestMethod("CreateFolder関数をユニットテストする")
Private Sub Test_CreateFolder()
    'CreateFolder(p)は引数として指定されたpをフォルダのパスとみなしてそのフォルダを作る。
    'pがまだ無ければ新しく作る。pがすでにあったらなにもしない。
    'pの親フォルダがまだ無ければエラー。親フォルダを作るにはEnsureFolder関数を使え。
    Dim p As String: p = ThisWorkbook.path & "\" & "tmp"
    CreateFolder (p)
    Assert.IsTrue PathExists(p)
    DeleteFolder (p)
End Sub

'@TestMethod("EnsureFolders関数をユニットテストする")
Private Sub Test_EnsureFolders()
    'EnsureFolders(p)はフォルダを作る。pの親フォルダが無かったらその祖先にまで遡って作る。
    Dim p As String: p = ThisWorkbook.path & "\build\tmp\testOutput"
    EnsureFolders (p)
    Assert.IsTrue PathExists(p)
    DeleteFolder (p)
End Sub

'@TestMethod("PathExists関数をユニットテストする")
Private Sub Test_PathExists()
    Assert.IsTrue PathExists(ThisWorkbook.path)
    Dim p As String: p = ThisWorkbook.path & "\" & "kazurayam-vba-lib.xlsm"
    Assert.IsTrue PathExists(p)
End Sub

'@TestMethod("WriteTextIntoFile関数とDeleteFile関数をテストする")
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

'@TestMethod("VerifyWorksheetExists(sheetName As String)をテストする")
Private Sub Test_VerifyWorksheetExists()
    'Assert:
    Assert.IsTrue VerifyWorksheetExists("Sheet1")
    Assert.IsFalse G.VerifyWorksheetExists("No Such Worksheet")
End Sub

'@TestMethod("DeleteWorksheetIfExists(sheetName As String)をテストする")
Private Sub Test_DeleteWorkSheetIfExists()
    'Arrange
    ' カレントのWorkbookにワークシートを挿入する、
    ' シートの名前はTest_DeleteWorkSheetIfExists
    Dim wsName As String: wsName = "Test_DeleteWorksheetIfExists"
    Worksheets.Add(After:=Worksheets(Worksheets.Count)) _
        .Name = wsName
    'Act:
    ' 挿入したワークシートを削除する
    G.DeleteWorksheetIfExists (wsName)
    ' 一時的に挿入したワークシートがもはや存在しないことを確認する
    Assert.IsFalse G.VerifyWorksheetExists(wsName)
End Sub

'@TestMethod("ExistsKey関数をユニットテストする")
Private Sub Test_ExistsKey()
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
    Assert.IsTrue ExistsKey(oCol, "炊飯器")
    Assert.IsFalse ExistsKey(oCol, "ルンバ")
End Sub


