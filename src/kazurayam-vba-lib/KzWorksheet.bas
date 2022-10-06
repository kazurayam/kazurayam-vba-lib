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

Public Function KzIfWorksheetExistsInWorkbook(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    Dim flg As Boolean: flg = False
    For Each ws In wb.Worksheets
        If ws.Name = sheetName Then
            flg = True
            Exit For
        End If
    Next ws
    KzIfWorksheetExistsInWorkbook = flg
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
' �������Ɍ��ƂȂ�Workbook�C���X�^���X���w�肷��B
' ��������String�ŁA�������Ŏw�肳�ꂽ���[�N�u�b�N�̂Ȃ��ɂ��郏�[�N�V�[�g�̖��O���w�肷��B
' �������Ƒ������ɂ��R�s�[��source�ƂȂ郏�[�N�V�[�g����肷��B
' ��O�����͐�ƂȂ�Worksheet�C���X�^���X���w�肷��B
' ��l������String�ŁA��̃��[�N�u�b�N�̂Ȃ��̃��[�N�V�[�g�����w�肷��B
' ��l�����Ƃ��Ďw�肳�ꂽ���O�̃��[�N�V�[�g����O�����̃��[�N�u�b�N�ɖ���������A���[�N�V�[�g��V�����}������B
' ��l�����Ƃ��Ďw�肳�ꂽ���O�̃��[�N�V�[�g����O�����̃��[�N�u�b�N�ɂ��łɑ��݂��Ă�����A
' ���̃��[�N�V�[�g�̂Ȃ��̃f�[�^��S���������Ă���Asource�̃��[�N�V�[�g����f�[�^����荞�ށB
' �Q�ƁFhttps://akira55.com/other_books/
Public Sub KzImportWorksheetFromWorkbook(ByVal sourceWorkbook As Workbook, _
                                        ByVal sourceSheetName As String, _
                                        ByVal targetWorkbook As Workbook, _
                                        ByVal targetSheetName As String)
    '�\��t����̃��[�N�V�[�g�����łɂ���������e������������
    If KzVerifyWorksheetExists(targetSheetName) Then
        targetWorkbook.Worksheets(targetSheetName).Cells.Clear
    Else
        '�\��t����̃��[�N�V�[�g���܂������������̃V�[�g��}������
        Dim ws As Worksheet
        Set ws = targetWorkbook.Worksheets.Add
        ws.Name = targetSheetName
    End If
    '�R�s�[�����[�N�V�[�g�̂��ׂẴZ�����R�s�[����
    sourceWorkbook.Worksheets(sourceSheetName).Cells.Copy
    '�R�s�[�惏�[�N�V�[�g�\��t����
    targetWorkbook.Worksheets(targetSheetName).Range("A1").PasteSpecial xlPasteFormulasAndNumberFormats
    Application.CutCopyMode = False '�R�s�[�؂��������
End Sub

