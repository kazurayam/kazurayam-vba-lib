Attribute VB_Name = "KzWorksheet"
Option Explicit

'KzWorksheet


' 指定された名のシートがカレントのブックに存在していたらTrueを返す
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

' 指定された名のシートがカレントのブックのなかに存在すれば削除する
Public Function KzDeleteWorksheetIfExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    Dim flg As Boolean: flg = False
    '指定されたブックに指定したシートが存在するかチェック
    For Each ws In Worksheets
        If ws.Name = sheetName Then
            'あればシートを削除する
            Application.DisplayAlerts = False    ' メッセージを非表示
            ws.Delete
            Application.DisplayAlerts = True
            flg = True
            Exit For
        End If
    Next ws
    KzDeleteWorksheetIfExists = flg
End Function


' コピー元として指定されたワークブックのワークシートをコピー先として指定されたワークブックのワークシートにコピーする。
' @param sourceWorkbook コピー元のWorkbook
' @param sourceSheetName コピー元のWorksheetの名前
' @param targetWorkbook コピー先のWorkbook
' @param targetSheetName コピー先のWorksheetの名前
' sourceWorkbookとsourceSheetNameで示されるワークシートが存在していることが必要。さもなければエラーになる。
' targetWorkbookのなかにtargetSheetNameで示されるワークシートが未だ無い場合とすでに在る場合とがありうる。
' まだ無ければsourceのシートをコピーするので、けっきょくtargetSheetNameのワークシートが新しくできる。
' すでに在ったら古いシートを削除してからsourceのシートをコピーする。
' ただしsourceWorkbookとtargetWorkbookが同じで、かつ、sourceSheetNameとtargetSheetNameが同じ場合は指定の誤りとみなしてエラーとする。
'
Public Sub KzImportWorksheetFromWorkbook(ByVal sourceWorkbook As Workbook, _
                                        ByVal sourceSheetName As String, _
                                        ByVal targetWorkbook As Workbook, _
                                        ByVal targetSheetName As String)
    'sourceとtargetが同じ場合はエラーとする
    If sourceWorkbook.path = targetWorkbook.path And sourceSheetName = targetSheetName Then
        Err.Raise Number:=2022, Description:="同じワークブック同じワークシートをsourceにtargetとして指定してはいけません"
    End If
    '貼り付け先のワークシートがすでにターゲットのワークブックのなかにあったら削除する
    If KzIfWorksheetExistsInWorkbook(targetWorkbook, targetSheetName) Then
        targetWorkbook.Worksheets(targetSheetName).Delete
    End If
    'コピー元ワークシートのすべてのセルをコピーしてターゲットのワークブックに新しいワークシートを挿入し
    sourceWorkbook.Worksheets(sourceSheetName).Copy After:=targetWorkbook.Sheets(targetWorkbook.Sheets.Count)
    '新しいワークシートの名前を指定されたようになおす
    ActiveSheet.Name = targetSheetName
End Sub

