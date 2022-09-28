Attribute VB_Name = "Test_G"
Option Explicit

' G.ToLocalFilePath()��P�̃e�X�g����
Sub Test_ToLocalFilePath()
    G.Cls
    Debug.Print vbCrLf; "---- Test_ToLocalFilePath() ----"
    Dim Source As String: Source = "https://d.docs.live.net/c5960fe753e170b9/�f�X�N�g�b�v/Excel-Word-VBA"
    Dim expect As String: expect = "C:\Users\uraya\OneDrive\�f�X�N�g�b�v\Excel-Word-VBA"
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

' G.VerifyWorksheetExists(sheetName As String)���e�X�g����
Sub Test_VerifyWorksheetExists()
    Debug.Assert G.VerifyWorksheetExists("Sheet1")
    Debug.Assert Not G.VerifyWorksheetExists("No Such Worksheet")
End Sub

' G.DeleteWorksheetIfExists(sheetName As String)���e�X�g����
Sub Test_DeleteWorkSheetIfExists()
    ' �J�����g��Workbook�Ƀ��[�N�V�[�g��}������A
    ' �V�[�g�̖��O��Test_DeleteWorkSheetIfExists
    Dim wsName As String: wsName = "Test_DeleteWorksheetIfExists"
    Worksheets.Add(After:=Worksheets(Worksheets.Count)) _
        .Name = wsName
    ' �}���������[�N�V�[�g���폜����
    G.DeleteWorksheetIfExists (wsName)
    ' �ꎞ�I�ɑ}���������[�N�V�[�g�����͂⑶�݂��Ȃ����Ƃ��m�F����
    Debug.Assert Not G.VerifyWorksheetExists(wsName)
End Sub

Sub Test_ExistsKey()
    Dim oCol As New Collection
    With oCol
        .Add Key:="�e���r", Item:="TV"
        .Add Key:="�①��", Item:="fridge"
        .Add Key:="���ъ�", Item:="rice cooker"
    End With
    Debug.Assert G.ExistsKey(oCol, "���ъ�")
    Debug.Assert Not G.ExistsKey(oCol, "�����o")
End Sub
