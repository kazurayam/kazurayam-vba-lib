Attribute VB_Name = "KzCollection"
Option Explicit

'KzCollection

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
 
Public Function KzExistsKey(objCol As Collection, strKey As String) As Boolean
     
    '�߂�l�̏����l�FFalse
    KzExistsKey = False
     
    '�ϐ���Collection���ݒ�̏ꍇ�͏����I��
    If objCol Is Nothing Then Exit Function
     
    'Collection�̃����o�[�����u0�v�̏ꍇ�͏����I��
    If objCol.Count = 0 Then Exit Function
     
    On Error Resume Next
     
    'Item���\�b�h�����s
    Call objCol.Item(strKey)
         
    '�G���[�l���Ȃ��ꍇ�F�L�[�����̓q�b�g�i�߂�l�FTrue�j
    If Err.Number = 0 Then KzExistsKey = True
 
End Function

