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


' �R�s�[���Ƃ��Ďw�肳�ꂽ���[�N�u�b�N�̃��[�N�V�[�g���R�s�[��Ƃ��Ďw�肳�ꂽ���[�N�u�b�N�̃��[�N�V�[�g�ɃR�s�[����B
' @param sourceWorkbook �R�s�[����Workbook
' @param sourceSheetName �R�s�[����Worksheet�̖��O
' @param targetWorkbook �R�s�[���Workbook
' @param targetSheetName �R�s�[���Worksheet�̖��O
' sourceWorkbook��sourceSheetName�Ŏ�����郏�[�N�V�[�g�����݂��Ă��邱�Ƃ��K�v�B�����Ȃ���΃G���[�ɂȂ�B
' targetWorkbook�̂Ȃ���targetSheetName�Ŏ�����郏�[�N�V�[�g�����������ꍇ�Ƃ��łɍ݂�ꍇ�Ƃ����肤��B
' �܂��������source�̃V�[�g���R�s�[����̂ŁA�������傭targetSheetName�̃��[�N�V�[�g���V�����ł���B
' ���łɍ݂�����Â��V�[�g���폜���Ă���source�̃V�[�g���R�s�[����B
' ������sourceWorkbook��targetWorkbook�������ŁA���AsourceSheetName��targetSheetName�������ꍇ�͎w��̌��Ƃ݂Ȃ��ăG���[�Ƃ���B
'
Public Sub KzImportWorksheetFromWorkbook(ByVal sourceWorkbook As Workbook, _
                                        ByVal sourceSheetName As String, _
                                        ByVal targetWorkbook As Workbook, _
                                        ByVal targetSheetName As String)
    'source��target�������ꍇ�̓G���[�Ƃ���
    If sourceWorkbook.path = targetWorkbook.path And sourceSheetName = targetSheetName Then
        Err.Raise Number:=2022, Description:="�������[�N�u�b�N�������[�N�V�[�g��source��target�Ƃ��Ďw�肵�Ă͂����܂���"
    End If
    '�\��t����̃��[�N�V�[�g�����łɃ^�[�Q�b�g�̃��[�N�u�b�N�̂Ȃ��ɂ�������폜����
    If KzIfWorksheetExistsInWorkbook(targetWorkbook, targetSheetName) Then
        targetWorkbook.Worksheets(targetSheetName).Delete
    End If
    '�R�s�[�����[�N�V�[�g�̂��ׂẴZ�����R�s�[���ă^�[�Q�b�g�̃��[�N�u�b�N�ɐV�������[�N�V�[�g��}����
    sourceWorkbook.Worksheets(sourceSheetName).Copy After:=targetWorkbook.Sheets(targetWorkbook.Sheets.Count)
    '�V�������[�N�V�[�g�̖��O���w�肳�ꂽ�悤�ɂȂ���
    ActiveSheet.Name = targetSheetName
End Sub

