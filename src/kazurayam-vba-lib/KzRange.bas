Attribute VB_Name = "KzRange"
Option Explicit

' KzRange���W���[���@Range�I�u�W�F�N�g�𑀍삷��w���p����������

'Range�I�u�W�F�N�g���󂯎��A���e�������z��ɕϊ����ĕԂ��B������
'�Ԃ����z��̓��e��String�ŁA�d���������B
Public Function KzGetUniqueItems(r As Range) As Variant
    Dim d As Dictionary: Set d = New Dictionary
    Dim item As Variant
    Dim k As String
    For Each item In r
        k = CStr(item)   '�L�[�𖾎��I��String�ɂ���
        If d.Exists(k) = False Then
            d.Add k, item
            'Debug.Print "add " & k
        End If
    Next
    KzGetUniqueItems = d.Keys
End Function

