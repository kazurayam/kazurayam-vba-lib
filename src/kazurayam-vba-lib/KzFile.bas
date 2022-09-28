Attribute VB_Name = "KzFile"
Option Explicit

'KzFile

Public Function KzAbsolutifyPath(ByVal basePath As String, ByVal RefPath As String) As String
    ' �t�@�C���̑��΃p�X���΃p�X�ɕϊ�����
    ' basePath�Ɋ����w�肷��
    Dim objFso As Object: Set objFso = CreateObject("Scripting.FileSystemObject")
    KzAbsolutifyPath = objFso.GetAbsolutePathName(objFso.BuildPath(basePath, RefPath))
    Set objFso = Nothing
End Function


' ����path�� "https://d.docs.live.net/c5960fe753e170b9/�f�X�N�g�b�v/Excel-Word-VBA" �̂悤��
' ���̃t�@�C����OneDrive�Ƀ}�b�s���O����Ă��邱�Ƃ�����URL�����񂩂ǂ����𒲂ׂ�B
' ���������Ȃ�� "C:\Users" �Ŏn�܂�OneDrive�̃��[�J���Ȍ`����String�ɏ��������ĕԂ��B
' ���������łȂ����path�����̂܂ܕԂ��B
Public Function KzToLocalFilePath(ByVal path As String) As String
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
        KzToLocalFilePath = p
        Set objFso = Nothing
    Else
        KzToLocalFilePath = path
    End If
End Function


' String�Ƃ��ăp�X���w�肳�ꂽ�t�H���_�����łɑ��݂��Ă��邩�ǂ����𒲂ׂ�
' ����������������������B�������e�t�H���_�������ꍇ�ɂ͎��s����B
Public Sub KzCreateFolder(folderPath As String)
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
Public Sub KzEnsureFolders(path As String)
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
Public Function KzPathExists(ByVal path As String) As Boolean
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim flg As Boolean: flg = False
    If fso.FileExists(path) Then
        flg = True
    ElseIf fso.FolderExists(path) Then
        flg = True
    End If
    KzPathExists = flg
End Function





' �p�X���w�肵���t�@�C�������݂��Ă�����폜����B
' �t�@�C����������΂Ȃɂ����Ȃ��B
Public Sub KzDeleteFile(ByVal fileToDelete As String)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(fileToDelete) Then 'See above
        ' First remove readonly attribute, if set
        SetAttr fileToDelete, vbNormal
        ' Then delete the file
        Kill fileToDelete
    End If
End Sub


' �t�H���_�����݂��Ă�����폜����
Public Sub KzDeleteFolder(ByVal folderToDelete As String)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(folderToDelete) Then
        fso.DeleteFolder (folderToDelete)
    End If
End Sub

' �e�L�X�g���t�@�C����WRITE����B
' �t�@�C����[�߂�ׂ��e�t�H���_��������΍���Ă���B
Public Sub KzWriteTextIntoFile(ByVal textData As String, ByVal file As String)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    KzEnsureFolders (fso.getParentFolderName(file))
    If fso.FileExists(file) Then
        KzDeleteFile (file)
    End If
    Dim fileNo As Integer
    fileNo = FreeFile
    Open file For Output As #fileNo
    Write #fileNo, textData
    Close #fileNo
End Sub



