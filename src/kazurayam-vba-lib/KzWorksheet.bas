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


' �R�s�[���̃��[�N�V�[�g�̃f�[�^�S�̂��R�s�[��̃��[�N�V�[�g�ɃR�s�[����B
' ��������Workbook�C���X�^���X���w�肷��B
' ��������String�ŁA�������Ŏw�肳�ꂽ���[�N�u�b�N�̂Ȃ��ɂ��郏�[�N�V�[�g�̖��O���w�肷��B
' �������Ƒ������ɂ��R�s�[��source�ƂȂ郏�[�N�V�[�g����肷��B
' ��O������String�ŁA�J�����g�̃��[�N�u�b�N�̂Ȃ��̃��[�N�V�[�g�����w�肷��B������R�s�[��Ɖ��߂���B�ȗ��ł��Ȃ��B
' ��O�����Ƃ��Ďw�肳�ꂽ���O�̃��[�N�V�[�g���J�����g�̃��[�N�u�b�N�ɖ���������A���[�N�V�[�g��V�����}������B
' ��O�����Ƃ��Ďw�肳�ꂽ���O�̃��[�N�V�[�g���J�����g�̃��[�N�u�b�N�ɂ��łɑ��݂��Ă�����A
' ���̃��[�N�V�[�g�̂Ȃ��̃f�[�^��S���������Ă���Asource�̃��[�N�V�[�g����f�[�^����荞�ށB
' �Q�ƁFhttps://akira55.com/other_books/
Public Sub KzImportWorksheetFromWorkbook(ByVal wbSource As Workbook, _
                                        ByVal sourceSheetName As String, _
                                        ByVal targetSheetName As String)
    Dim wbTarget As Workbook: Set wbTarget = ActiveWorkbook
    '�\��t����̃��[�N�V�[�g�����łɂ���������e������������
    If KzVerifyWorksheetExists(targetSheetName) Then
        wbTarget.Worksheets(targetSheetName).Cells.Clear
    Else
        '�\��t����̃��[�N�V�[�g���܂������������̃V�[�g��}������
        wbTarget.Worksheets.Add
        ActiveSheet.Name = targetSheetName
    End If
    '�R�s�[�����[�N�V�[�g�̂��ׂẴZ�����R�s�[����
    wbSource.Worksheets(sourceSheetName).Cells.Copy
    '�R�s�[�惏�[�N�V�[�g�\��t����
    wbTarget.Worksheets(targetSheetName).Range("A1").PasteSpecial xlPasteFormulasAndNumberFormats
    Application.CutCopyMode = False '�R�s�[�؂��������
End Sub

