Attribute VB_Name = "KzUtil"
Option Explicit

'KzUtil

' Clear Immediate Window
' calls Debug.Print many times to output blank lines
' so that the immediate window is wiped out
Public Sub KzCls()
    Debug.Print String(200, vbCrLf)
End Sub


Public Function KzVarTypeAsString(ByVal var As Variant) As String
    ' ����var��type�𒲂ׂĕϐ��̌^������������i"Integer"�Ȃǁj��Ԃ�
    Dim typeValue As Long: typeValue = VarType(var)
    Dim result As String: result = "unknown"
    If typeValue = 2 Then
        result = "Integer"
    ElseIf typeValue = 3 Then
        result = "Long"
    ElseIf typeValue = 5 Then
        result = "Double"
    ElseIf typeValue = 8 Then
        result = "String"
    ElseIf typeValue = 11 Then
        result = "Boolean"
    ElseIf typeValue = 7 Then
        result = "Date"
    ElseIf typeValue = 9 Then
        result = "Object"
    ElseIf typeValue = 0 Then
        result = "Variant"
    ElseIf typeValue = 8200 Then
        result = "String()"
    ElseIf typeValue = 8194 Then
        result = "Integer()"
    Else
        result = Str(typeValue)
    End If
    KzVarTypeAsString = result
End Function


Public Function KzResolveExternalFilePath( _
        ByVal theWorkbook As Workbook, _
        ByVal sheetName As String, _
        ByVal rangeLiteral As String) As String
    'theWorkbook�Ƃ��ė^����ꂽ���[�N�u�b�N�̂Ȃ���
    'sheetName�Ƃ��ė^����ꂽ���[�N�V�[�g�������āA���̒���
    'rangeLiteral�Ƃ��ė^����ꂽ�Z���������āA���̂Ȃ���
    '�O���t�@�C���̃p�X�������Ă���Ɗ��҂���B
    '���̃p�X��theWorkbook�����Ƃ��鑊�΃p�X�ł���Ɗ��҂���B
    '�O���t�@�C���̃p�X�𔭌����A������΃p�X�ɕϊ����āAFunction�̒l�Ƃ��ĕԂ��B
    '���̊֐���.xlsm�t�@�C���̉��������߂�̂ɗL�p�ł���B
    '.xlsm�t�@�C�����猩���O���t�@�C���̃p�X��VBA�R�[�h�̂Ȃ���
    '�Œ�l�Ƃ��ď����̂ł͂Ȃ��A
    '���[�N�V�[�g�̃Z���̒l�Ƃ��ď������Ƃ��\�ɂ���B
    Dim ws As Worksheet: Set ws = theWorkbook.Worksheets(sheetName)
    Dim path As String
    path = ws.Range(rangeLiteral)
    
    KzResolveExternalFilePath = KzAbsolutifyPath(KzToLocalFilePath(theWorkbook.path), path)
End Function


'KzResolveExternalFilePath�֐����e�X�g����
Private Sub Test_KzResolveExternalFilePath()
    Dim p As String
    p = KzResolveExternalFilePath(ThisWorkbook, "�O�����[�N�u�b�N�t�@�C���̃p�X", "B2")
    Debug.Print p
End Sub
