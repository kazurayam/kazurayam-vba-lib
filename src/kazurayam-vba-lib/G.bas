Attribute VB_Name = "G"
Option Explicit

' G: A collection of Global Public Functions that is applicable to any type of VBA projects

' Clear Immediate Window
' calls Debug.Print many times to output blank lines
' so that the immediate window is wiped out
Public Sub Cls()
    Dim i As Long
    For i = 0 To 20
        Debug.Print
    Next i
End Sub


Function AbsolutifyPath(ByVal basePath As String, ByVal RefPath As String) As String
    ' ファイルの相対パスを絶対パスに変換する
    ' basePathに基底を指定する
    Dim objFso As Object: Set objFso = CreateObject("Scripting.FileSystemObject")
    AbsolutifyPath = objFso.GetAbsolutePathName(objFso.BuildPath(basePath, RefPath))
    Set objFso = Nothing
End Function


' 引数pathが "https://d.docs.live.net/c5960fe753e170b9/デスクトップ/Excel-Word-VBA" のように
' そのファイルがOneDriveにマッピングされていることを示すURL文字列かどうかを調べる。
' もしそうならば "C:\Users" で始まるOneDriveのローカルな形式のStringに書きかえて返す。
' もしそうでなければpathをそのまま返す。
Function ToLocalFilePath(ByVal path As String) As String
    Dim searchResult As Integer
    searchResult = VBA.Strings.InStr(1, path, "https://d.docs.live.net/", vbTextCompare)
    ' Debug.Print "searchResult=" & searchResult
    If searchResult = 1 Then
        Dim s() As String
        s = VBA.Strings.Split(path, "/", Limit:=5, Compare:=vbBinaryCompare)
        ' sは配列で中身は Arrays("https:", "", "d.docs.live.net", "c5960fe753e170b9", "デスクトップ/Excel-Word-VBA") になっている
        Dim objFso As Object
        Set objFso = CreateObject("Scripting.FileSystemObject")
        Dim p As String: p = objFso.GetAbsolutePathName(objFso.BuildPath(VBA.Interaction.Environ("OneDrive"), s(UBound(s))))
        ' UBound関数は引数に指定した配列で使用できる最も大きいインデックス番号を返す
        ' s(UBound(s)) は 配列sの5番目の要素 "デスクトップ/Excel-Word-VBA" を返す
        ' VBA.Interaction.Environ()は環境変数の値を返す
        ' Environ("OneDrive")の値はたとえば "C:\Users\uraya\OneDrive" という文字列を返す
        ' objFso.BuildPath(path, name)はパスとファイル名を表す二つの文字列を連結してひとつの文字列を返す。/は\に置き換えられる。
        ' objFso.GetAbsolutePathName(pathspec)は pathspec（相対パスかもしれない）を絶対パスに変換します
        ToLocalFilePath = p
        Set objFso = Nothing
    Else
        ToLocalFilePath = path
    End If
End Function


' Stringとしてパスが指定されたフォルダがすでに存在しているかどうかを調べて
' もしも未だ無かったら作る。ただし親フォルダも無い場合には失敗する。
Sub CreateFolder(folderPath As String)
    Dim objFso As Object
    Set objFso = CreateObject("Scripting.FileSystemObject")
    If objFso.FolderExists(folderPath) Then
        ' does nothing
    Else
        objFso.CreateFolder (folderPath)
        Debug.Print "created " & folderPath
    End If
    Set objFso = Nothing
End Sub

' フォルダのフルパスが与えられることを前提する。フォルダを作る。
' ルートから子孫フォルダを順番に有無をしらべて、無ければMkDirで作る。
' つまり指定されたフォルダの先祖が無ければ先祖も作ってしまう。
Sub EnsureFolders(path As String)
    Dim tmp As String
    Dim arr() As String
    arr = Split(path, "\")
    tmp = arr(0)
    Dim i As Long
    For i = LBound(arr) + 1 To UBound(arr)
        tmp = tmp & "\" & arr(i)
        If Dir(tmp, vbDirectory) = "" Then
            ' フォルダが無ければ作る
            MkDir tmp
        End If
    Next i
End Sub

' pathが示すパスにファイルまたはフォルダが存在していたらTrueをかえす。
' pathが示すパスにファイルもフォルダも無いならFalseをかえす
Function PathExists(ByVal path As String) As Boolean
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim flg As Boolean: flg = False
    If fso.FileExists(path) Then
        flg = True
    ElseIf fso.FolderExists(path) Then
        flg = True
    End If
    PathExists = flg
End Function

' パスを指定したファイルが存在していたら削除する。
' ファイルが無ければなにもしない。
Sub DeleteFile(ByVal fileToDelete As String)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(fileToDelete) Then 'See above
        ' First remove readonly attribute, if set
        SetAttr fileToDelete, vbNormal
        ' Then delete the file
        Kill fileToDelete
    End If
End Sub

' フォルダが存在していたら削除する
Sub DeleteFolder(ByVal folderToDelete As String)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(folderToDelete) Then
        fso.DeleteFolder (folderToDelete)
    End If
End Sub

' テキストをファイルにWRITEする。
' ファイルを納めるべき親フォルダが無ければ作ってから。
Sub WriteTextIntoFile(ByVal textData As String, ByVal file As String)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    G.EnsureFolders (fso.getParentFolderName(file))
    If fso.FileExists(file) Then
        G.DeleteFile (file)
    End If
    Dim fileNo As Integer
    fileNo = FreeFile
    Open file For Output As #fileNo
    Write #fileNo, textData
    Close #fileNo
End Sub



' 指定された名のシートがカレントのブックに存在していたらTrueを返す
Public Function VerifyWorksheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    Dim flg As Boolean: flg = False
    For Each ws In Worksheets
        If ws.Name = sheetName Then
            flg = True
            Exit For
        End If
    Next ws
    VerifyWorksheetExists = flg
End Function

' 指定された名のシートがカレントのブックのなかに存在すれば削除する
Public Function DeleteWorksheetIfExists(sheetName As String) As Boolean
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
    DeleteWorksheetIfExists = flg
End Function

' https://y-moride.com/vba/collection-key-exists.html#toc1
'*********************************************************
'* ExistsKey（Collention内のキー検索関数）
'*********************************************************
'* 第１引数 | Collection | 検索対象となるオブジェクト
'* 第２引数 |   String   | 検索するキー
'*  戻り値　|   Boolan   | True Or False ※False＠初期値
'*********************************************************
'*   説明   | 第２引数をキーとしてItemメソッドを実行し、
'*   　　   | 結果をもとにキーの存在を確認する。
'*********************************************************
'*   備考   | オブジェクト未設定の場合 ⇒ 戻り値「False」
'*   　　   | メンバー数「0」の場合 ⇒ 戻り値「False」
'*********************************************************
 
Function ExistsKey(objCol As Collection, strKey As String) As Boolean
     
    '戻り値の初期値：False
    ExistsKey = False
     
    '変数にCollection未設定の場合は処理終了
    If objCol Is Nothing Then Exit Function
     
    'Collectionのメンバー数が「0」の場合は処理終了
    If objCol.Count = 0 Then Exit Function
     
    On Error Resume Next
     
    'Itemメソッドを実行
    Call objCol.Item(strKey)
         
    'エラー値がない場合：キー検索はヒット（戻り値：True）
    If Err.Number = 0 Then ExistsKey = True
 
End Function
