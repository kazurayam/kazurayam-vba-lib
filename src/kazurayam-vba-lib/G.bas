Attribute VB_Name = "G"
Option Explicit

' G: A collection of Global Public Functions that is applicable to any type of VBA projects

' Clear Immediate Window
' calls Debug.Print many times to output blank lines
' so that the immediate window is wiped out
Public Sub Cls()
    Dim i As Long
    For i = 0 To 20
        Debug.Print
    Next i
End Sub


Function AbsolutifyPath(ByVal basePath As String, ByVal RefPath As String) As String
    ' �t�@�C���̑��΃p�X���΃p�X�ɕϊ�����
    ' basePath�Ɋ����w�肷��
    Dim objFso As Object: Set objFso = CreateObject("Scripting.FileSystemObject")
    AbsolutifyPath = objFso.GetAbsolutePathName(objFso.BuildPath(basePath, RefPath))
    Set objFso = Nothing
End Function


' ����path�� "https://d.docs.live.net/c5960fe753e170b9/�f�X�N�g�b�v/Excel-Word-VBA" �̂悤��
' ���̃t�@�C����OneDrive�Ƀ}�b�s���O����Ă��邱�Ƃ�����URL�����񂩂ǂ����𒲂ׂ�B
' ���������Ȃ�� "C:\Users" �Ŏn�܂�OneDrive�̃��[�J���Ȍ`����String�ɏ��������ĕԂ��B
' ���������łȂ����path�����̂܂ܕԂ��B
Function ToLocalFilePath(ByVal path As String) As String
    Dim searchResult As Integer
    searchResult = VBA.Strings.InStr(1, path, "https://d.docs.live.net/", vbTextCompare)
    ' Debug.Print "searchResult=" & searchResult
    If searchResult = 1 Then
        Dim s() As String
        s = VBA.Strings.Split(path, "/", Limit:=5, Compare:=vbBinaryCompare)
        ' s�͔z��Œ��g�� Arrays("https:", "", "d.docs.live.net", "c5960fe753e170b9", "�f�X�N�g�b�v/Excel-Word-VBA") �ɂȂ��Ă���
        Dim objFso As Object
        Set objFso = CreateObject("Scripting.FileSystemObject")
        Dim p As String: p = objFso.GetAbsolutePathName(objFso.BuildPath(VBA.Interaction.Environ("OneDrive"), s(UBound(s))))
        ' UBound�֐��͈����Ɏw�肵���z��Ŏg�p�ł���ł��傫���C���f�b�N�X�ԍ���Ԃ�
        ' s(UBound(s)) �� �z��s��5�Ԗڂ̗v�f "�f�X�N�g�b�v/Excel-Word-VBA" ��Ԃ�
        ' VBA.Interaction.Environ()�͊��ϐ��̒l��Ԃ�
        ' Environ("OneDrive")�̒l�͂��Ƃ��� "C:\Users\uraya\OneDrive" �Ƃ����������Ԃ�
        ' objFso.BuildPath(path, name)�̓p�X�ƃt�@�C������\����̕������A�����ĂЂƂ̕������Ԃ��B/��\�ɒu����������B
        ' objFso.GetAbsolutePathName(pathspec)�� pathspec�i���΃p�X��������Ȃ��j���΃p�X�ɕϊ����܂�
        ToLocalFilePath = p
        Set objFso = Nothing
    Else
        ToLocalFilePath = path
    End If
End Function


' String�Ƃ��ăp�X���w�肳�ꂽ�t�H���_�����łɑ��݂��Ă��邩�ǂ����𒲂ׂ�
' ����������������������B�������e�t�H���_�������ꍇ�ɂ͎��s����B
Sub CreateFolder(folderPath As String)
    Dim objFso As Object
    Set objFso = CreateObject("Scripting.FileSystemObject")
    If objFso.FolderExists(folderPath) Then
        ' does nothing
    Else
        objFso.CreateFolder (folderPath)
        Debug.Print "created " & folderPath
    End If
    Set objFso = Nothing
End Sub

' �t�H���_�̃t���p�X���^�����邱�Ƃ�O�񂷂�B�t�H���_�����B
' ���[�g����q���t�H���_�����ԂɗL��������ׂāA�������MkDir�ō��B
' �܂�w�肳�ꂽ�t�H���_�̐�c��������ΐ�c������Ă��܂��B
Sub EnsureFolders(path As String)
    Dim tmp As String
    Dim arr() As String
    arr = Split(path, "\")
    tmp = arr(0)
    Dim i As Long
    For i = LBound(arr) + 1 To UBound(arr)
        tmp = tmp & "\" & arr(i)
        If Dir(tmp, vbDirectory) = "" Then
            ' �t�H���_��������΍��
            MkDir tmp
        End If
    Next i
End Sub

' path�������p�X�Ƀt�@�C���܂��̓t�H���_�����݂��Ă�����True���������B
' path�������p�X�Ƀt�@�C�����t�H���_�������Ȃ�False��������
Function PathExists(ByVal path As String) As Boolean
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim flg As Boolean: flg = False
    If fso.FileExists(path) Then
        flg = True
    ElseIf fso.FolderExists(path) Then
        flg = True
    End If
    PathExists = flg
End Function

' �p�X���w�肵���t�@�C�������݂��Ă�����폜����B
' �t�@�C����������΂Ȃɂ����Ȃ��B
Sub DeleteFile(ByVal fileToDelete As String)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(fileToDelete) Then 'See above
        ' First remove readonly attribute, if set
        SetAttr fileToDelete, vbNormal
        ' Then delete the file
        Kill fileToDelete
    End If
End Sub

' �t�H���_�����݂��Ă�����폜����
Sub DeleteFolder(ByVal folderToDelete As String)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(folderToDelete) Then
        fso.DeleteFolder (folderToDelete)
    End If
End Sub

' �e�L�X�g���t�@�C����WRITE����B
' �t�@�C����[�߂�ׂ��e�t�H���_��������΍���Ă���B
Sub WriteTextIntoFile(ByVal textData As String, ByVal file As String)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    G.EnsureFolders (fso.getParentFolderName(file))
    If fso.FileExists(file) Then
        G.DeleteFile (file)
    End If
    Dim fileNo As Integer
    fileNo = FreeFile
    Open file For Output As #fileNo
    Write #fileNo, textData
    Close #fileNo
End Sub



' �w�肳�ꂽ���̃V�[�g���J�����g�̃u�b�N�ɑ��݂��Ă�����True��Ԃ�
Public Function VerifyWorksheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    Dim flg As Boolean: flg = False
    For Each ws In Worksheets
        If ws.Name = sheetName Then
            flg = True
            Exit For
        End If
    Next ws
    VerifyWorksheetExists = flg
End Function

' �w�肳�ꂽ���̃V�[�g���J�����g�̃u�b�N�̂Ȃ��ɑ��݂���΍폜����
Public Function DeleteWorksheetIfExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    Dim flg As Boolean: flg = False
    '�w�肳�ꂽ�u�b�N�Ɏw�肵���V�[�g�����݂��邩�`�F�b�N
    For Each ws In Worksheets
        If ws.Name = sheetName Then
            '����΃V�[�g���폜����
            Application.DisplayAlerts = False    ' ���b�Z�[�W���\��
            ws.Delete
            Application.DisplayAlerts = True
            flg = True
            Exit For
        End If
    Next ws
    DeleteWorksheetIfExists = flg
End Function

' https://y-moride.com/vba/collection-key-exists.html#toc1
'*********************************************************
'* ExistsKey�iCollention���̃L�[�����֐��j
'*********************************************************
'* ��P���� | Collection | �����ΏۂƂȂ�I�u�W�F�N�g
'* ��Q���� |   String   | ��������L�[
'*  �߂�l�@|   Boolan   | True Or False ��False�������l
'*********************************************************
'*   ����   | ��Q�������L�[�Ƃ���Item���\�b�h�����s���A
'*   �@�@   | ���ʂ����ƂɃL�[�̑��݂��m�F����B
'*********************************************************
'*   ���l   | �I�u�W�F�N�g���ݒ�̏ꍇ �� �߂�l�uFalse�v
'*   �@�@   | �����o�[���u0�v�̏ꍇ �� �߂�l�uFalse�v
'*********************************************************
 
Function ExistsKey(objCol As Collection, strKey As String) As Boolean
     
    '�߂�l�̏����l�FFalse
    ExistsKey = False
     
    '�ϐ���Collection���ݒ�̏ꍇ�͏����I��
    If objCol Is Nothing Then Exit Function
     
    'Collection�̃����o�[�����u0�v�̏ꍇ�͏����I��
    If objCol.Count = 0 Then Exit Function
     
    On Error Resume Next
     
    'Item���\�b�h�����s
    Call objCol.Item(strKey)
         
    '�G���[�l���Ȃ��ꍇ�F�L�[�����̓q�b�g�i�߂�l�FTrue�j
    If Err.Number = 0 Then ExistsKey = True
 
End Function
