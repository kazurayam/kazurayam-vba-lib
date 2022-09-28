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

