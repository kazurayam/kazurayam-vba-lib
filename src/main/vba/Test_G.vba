Attribute VB_Name = "Test_G"
Option Explicit

' G.ToLocalFilePath()を単体テストする
Sub Test_ToLocalFilePath()
    G.Cls
    Debug.Print vbCrLf; "---- Test_ToLocalFilePath() ----"
    Dim Source As String: Source = "https://d.docs.live.net/c5960fe753e170b9/デスクトップ/Excel-Word-VBA"
    Dim expect As String: expect = "C:\Users\uraya\OneDrive\デスクトップ\Excel-Word-VBA"
    Dim actual As String
    actual = ToLocalFilePath(Source)
    Debug.Print "source:" & vbTab; Chr(34); Source; Chr(34)
    Debug.Print "expect:" & vbTab; Chr(34); expect; Chr(34)
    Debug.Print "actual:" & vbTab; Chr(34); actual; Chr(34);
    Debug.Assert Len(actual) > 0
    Debug.Assert StrComp(expect, actual) = 0
End Sub

Sub Test_CreateFolder()
    Dim p As String: p = ThisWorkbook.path & "\" & "tmp"
    G.CreateFolder (p)
    Debug.Assert G.PathExists(p)
    G.DeleteFolder (p)
End Sub

Sub Test_EnsureFolders()
    Dim p As String: p = ThisWorkbook.path & "\build\tmp\testOutput"
    G.EnsureFolders (p)
    Debug.Assert G.PathExists(p)
    G.DeleteFolder (p)
End Sub

Sub Test_PathExists()
    Debug.Assert G.PathExists(ThisWorkbook.path)
    Dim p As String: p = ThisWorkbook.path & "\" & "kazurayam-vba-lib.xlsm"
    Debug.Assert G.PathExists(p)
End Sub

Sub Test_WriteTextIntoFile_and_DeleteFile()
    Dim folder As String: folder = ThisWorkbook.path & "\build"
    Dim file As String: file = folder & "\hello.txt"
    Call G.WriteTextIntoFile("Hello, world", file)
    Debug.Assert G.PathExists(file)
    G.DeleteFile (file)
End Sub

' G.VerifyWorksheetExists(sheetName As String)をテストする
Sub Test_VerifyWorksheetExists()
    Debug.Assert G.VerifyWorksheetExists("Sheet1")
    Debug.Assert Not G.VerifyWorksheetExists("No Such Worksheet")
End Sub

' G.DeleteWorksheetIfExists(sheetName As String)をテストする
Sub Test_DeleteWorkSheetIfExists()
    ' カレントのWorkbookにワークシートを挿入する、
    ' シートの名前はTest_DeleteWorkSheetIfExists
    Dim wsName As String: wsName = "Test_DeleteWorksheetIfExists"
    Worksheets.Add(After:=Worksheets(Worksheets.Count)) _
        .Name = wsName
    ' 挿入したワークシートを削除する
    G.DeleteWorksheetIfExists (wsName)
    ' 一時的に挿入したワークシートがもはや存在しないことを確認する
    Debug.Assert Not G.VerifyWorksheetExists(wsName)
End Sub

Sub Test_ExistsKey()
    Dim oCol As New Collection
    With oCol
        .Add Key:="テレビ", Item:="TV"
        .Add Key:="冷蔵庫", Item:="fridge"
        .Add Key:="炊飯器", Item:="rice cooker"
    End With
    Debug.Assert G.ExistsKey(oCol, "炊飯器")
    Debug.Assert Not G.ExistsKey(oCol, "ルンバ")
End Sub
