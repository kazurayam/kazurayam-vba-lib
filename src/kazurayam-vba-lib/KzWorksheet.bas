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
' 第一引数にWorkbookインスタンスを指定する。
' 第二引数はStringで、第一引数で指定されたワークブックのなかにあるワークシートの名前を指定する。
' 第一引数と第二引数によりコピー元sourceとなるワークシートを特定する。
' 第三引数はStringで、カレントのワークブックのなかのワークシート名を指定する。これをコピー先と解釈する。省略できない。
' 第三引数として指定された名前のワークシートがカレントのワークブックに無かったら、ワークシートを新しく挿入する。
' 第三引数として指定された名前のワークシートがカレントのワークブックにすでに存在していたら、
' そのワークシートのなかのデータを全部消去してから、sourceのワークシートからデータを取り込む。
' 参照：https://akira55.com/other_books/
Public Sub KzImportWorksheetFromWorkbook(ByVal wbSource As Workbook, _
                                        ByVal sourceSheetName As String, _
                                        ByVal targetSheetName As String)
    Dim wbTarget As Workbook: Set wbTarget = ActiveWorkbook
    '貼り付け先のワークシートがすでにあったら内容を初期化する
    If KzVerifyWorksheetExists(targetSheetName) Then
        wbTarget.Worksheets(targetSheetName).Cells.Clear
    Else
        '貼り付け先のワークシートがまだ無かったら空のシートを挿入する
        wbTarget.Worksheets.Add
        ActiveSheet.Name = targetSheetName
    End If
    'コピー元ワークシートのすべてのセルをコピーして
    wbSource.Worksheets(sourceSheetName).Cells.Copy
    'コピー先ワークシート貼り付ける
    wbTarget.Worksheets(targetSheetName).Range("A1").PasteSpecial xlPasteFormulasAndNumberFormats
    Application.CutCopyMode = False 'コピー切り取りを解除
End Sub

