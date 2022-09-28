Attribute VB_Name = "KzWorksheet"
Option Explicit

'KzWorksheet


' �w�肳�ꂽ���̃V�[�g���J�����g�̃u�b�N�ɑ��݂��Ă�����True��Ԃ�
Public Function KzVerifyWorksheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    Dim flg As Boolean: flg = False
    For Each ws In Worksheets
        If ws.Name = sheetName Then
            flg = True
            Exit For
        End If
    Next ws
    KzVerifyWorksheetExists = flg
End Function

' �w�肳�ꂽ���̃V�[�g���J�����g�̃u�b�N�̂Ȃ��ɑ��݂���΍폜����
Public Function KzDeleteWorksheetIfExists(sheetName As String) As Boolean
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
    KzDeleteWorksheetIfExists = flg
End Function

