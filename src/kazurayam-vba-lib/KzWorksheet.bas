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


' コピー元のワークシートのデータ全体をコピー先のワークシートにコピーする。
' 第一引数に元となるWorkbookインスタンスを指定する。
' 第二引数はStringで、第一引数で指定されたワークブックのなかにあるワークシートの名前を指定する。
' 第一引数と第二引数によりコピー元sourceとなるワークシートを特定する。
' 第三引数は先となるWorksheetインスタンスを指定する。
' 第四引数はStringで、先のワークブックのなかのワークシート名を指定する。
' 第四引数として指定された名前のワークシートが第三引数のワークブックに無かったら、ワークシートを新しく挿入する。
' 第四引数として指定された名前のワークシートが第三引数のワークブックにすでに存在していたら、
' そのワークシートのなかのデータを全部消去してから、sourceのワークシートからデータを取り込む。
' 参照：https://akira55.com/other_books/
Public Sub KzImportWorksheetFromWorkbook(ByVal sourceWorkbook As Workbook, _
                                        ByVal sourceSheetName As String, _
                                        ByVal targetWorkbook As Workbook, _
                                        ByVal targetSheetName As String)
    '貼り付け先のワークシートがすでにあったら内容を初期化する
    If KzVerifyWorksheetExists(targetSheetName) Then
        targetWorkbook.Worksheets(targetSheetName).Cells.Clear
    Else
        '貼り付け先のワークシートがまだ無かったら空のシートを挿入する
        Dim ws As Worksheet
        Set ws = targetWorkbook.Worksheets.Add
        ws.Name = targetSheetName
    End If
    'コピー元ワークシートのすべてのセルをコピーして
    sourceWorkbook.Worksheets(sourceSheetName).Cells.Copy
    'コピー先ワークシート貼り付ける
    targetWorkbook.Worksheets(targetSheetName).Range("A1").PasteSpecial xlPasteFormulasAndNumberFormats
    Application.CutCopyMode = False 'コピー切り取りを解除
End Sub

