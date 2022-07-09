# AddInモジュールのバージョン
$targetVersion = "0.1.0"

# AddInモジュールのURL on GitHub
$targetUrl = "https://github.com/kazurayam/kazurayam-vba-lib/raw/${targetVersion}/kazurayam-vba-lib.xlam"

# 配布先PCのOSユーザ名 
$userName = $env:UserName

# 配布先PC上でAddInモジュールを格納すべき宛先としてのフォルダ
$AddInsDir = "C:\\Users\\${userName}\\AppData\\Roaming\\Microsoft\\AddIns"

# WebClient 生成
$cli = New-Object System.Net.WebClient

# 対象URL
$uri = New-Object System.Uri($targetUrl)

# 保存時のファイル名を取得
$file = Split-Path $uri.AbsolutePath -Leaf

# ダウンロード
$cli.DownloadFile($uri, (Join-Path $AddInsDir $file))