# kazurayam-vba-lib

A Library of VBA procedures created by kazurayam for his own use

kazurayamが自分が使うために作ったVBAアドインのライブラリ。

## AddInにかんする技術情報

- https://tonari-it.com/excel-vba-class-addin/
- https://tonari-it.com/excel-vba-class-addin-reference/
- https://excel-ubara.com/excelvba4/EXCEL297.html


## xlsmファイルからxlamファイルを名前を付けて保存する

TODO

## AddInを配置する手順

### 開発者のための手順

kazurayamがこのGitレポジトリを自分のPCにcloneしたうえで、AddInを配置するには次の手順をとる。

1. `kazurayam-vba-lib.xlsm` ファイルをExcelで開く。
2. `名前を付けて保存`を選択する。出力するファイルの形式として `xlam` を選択する。このとき保存先フォルダの候補として `C:\Users\ユーザ名\AppData\Roaming\Microsoft\AddIns` フォルダが自動的に選択される。
3. 第一の方法はxlamファイルをExcelからAddInsフォルダの中に直接保存すること。
4. 第二に方法はxlamファイルをレポジトリのルートディレクトリの直下に保存したうえで、
```
$ cd kazurayam-vba-lib
$ ./gradle deployAddIn
```
とやること。Gradleスクリプトがレポジトリのルートディレクトリの直下にあるxlamファイルを`C:\Users\ユーザ名\AppData\Roaming\Microsoft\AddIns` フォルダにコピーする。
5. けっきょく `C:\Users\ユーザ名\AppData\Roaming\Microsoft\AddIns\kazurayam-vba-lib.xlam` ファイルができる。すでに同名のファイルが存在していたら上書き更新される。

### 利用者のための手順

kazurayam以外の一般の人がkazurayam-vba-libを自分のWindows PCに組み込んで使うこともできる。一般が自分のPCにAddInを配置するには次の手順をとる。

1. PowerShellスクリプト[updateAddIn.ps1](https://github.com/kazurayam/kazurayam-vba-lib/blob/master/updateAddIn.ps1)をテキストエディタで書いてデスクトップに保存する。
2. ダブルクリックして実行する。するとxlamファイルが https://github.com/kazurayam/kazurayam-vba-lib/blob/master/kazurayam-vba-lib.xlam からダウンロードされて `C:\Users\ユーザ名\AppData\Roaming\Microsoft\AddIns\kazurayam-vba-lib.xlam` にコピーされる。
3. 利用者はkazurayam-vba-libを自分のワークブックが参照するよう、ワークブックを設定しなければならない。すなわち 
  a. VBEのツールバーで　ツール＞参照設定　を選ぶ。参照設定のダイアログが表示される。
  b. 参照設定のダイアログの右側にあるボタンのうち　参照　ボタンをクリックする。
  c. `C:\Users\ユーザ名\AppData\Roaming\Microsoft\AddIns\kazurayam-vba-lib.xlam`　を選択する。OKボタンを押下する。

## AddInを開発するうえでの注意事項

1. `kazurayam-vba-lib.xlsm`ファイルを開いてVBEを開くとプロジェクト名を確認できる。プロジェクト名はユニークな名前にしなければならない。ここでは `KazurayamVbaLib`  とした
2. デフォルトとして`VBAProject`という名前が割り当てられるが、他のブックファイルでも同じプロジェクトが割り当てられるかもしれない。`VBAProject`が複数あると名前の重複のせいでプロシージャが見つからないというエラーになる。
