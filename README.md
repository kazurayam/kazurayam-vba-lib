# kazurayam-vba-lib

## これは何か

A Library of VBA procedures created by kazurayam for his own use。

kazurayamが自分が使うために作ったExce VBAアドインのライブラリ。


## User Guide 利用者向けガイド

kazurayamがバイト先で事務作業用のPCで仕事するときどうするかを説明しよう。またkazurayam以外の人がkazurayam-vba-libを自分のWindows PCでExcelからこのライブラリを参照して使ってもかまわない。どちらの場合も、ターゲットのPCでこのAddInを配置するには次の手順をとる。

1. [updateAddIn.ps1](https://github.com/kazurayam/kazurayam-vba-lib/blob/master/updateAddIn.ps1)ファイルをテキストエディタで書いてデスクトップに保存する。このスクリプトはPowerShellで書かれている。Windows10ならデフォルトでPowerShellが組み込まれている。PowerShellをインストールするなどの作業は不要。
2. updateAddIn.ps1を実行する。デスクトップアイコンをダブルクリックすればいい。https://github.com/kazurayam/kazurayam-vba-lib/blob/master/kazurayam-vba-lib.xlam からアドインのファイルがダウンロードされる。そして `C:\Users\ユーザ名\AppData\Roaming\Microsoft\AddIns\kazurayam-vba-lib.xlam` に配置される。
3. 利用者は 自分のExcelワークブックが kazurayam-vba-lib.xlam を参照するよう設定しなければならない。すなわち 
  a. VBEのツールバーで　ツール＞参照設定　を選ぶ。参照設定のダイアログが表示される。
  b. 参照設定のダイアログの右側にあるボタンのうち　参照　ボタンをクリックする。
  c. `C:\Users\ユーザ名\AppData\Roaming\Microsoft\AddIns\kazurayam-vba-lib.xlam`　を選択する。OKボタンを押下する。


## Developer Guide 開発者向けガイド

kazurayamが自宅で開発用のPCでこのライブラリを更新するときどうするか？

1. `kazurayam-vba-lib.xlsm` ファイルをExcelで開く。
2. `名前を付けて保存`を選択する。出力するファイルの形式として `xlam` を選択する。このとき保存先フォルダの候補として `C:\Users\ユーザ名\AppData\Roaming\Microsoft\AddIns` フォルダが自動的に選択される。
3. 第一の方法はxlamファイルをExcelからAddInsフォルダの中に保存する。ただしこの方法をとるとGitRepositoryにはxlamファイルが保存されないことに注意せよ。
4. 第二の方法はxlamファイルをレポジトリのルートディレクトリの直下に保存する。そのあと、
```
$ cd kazurayam-vba-lib
$ ./gradle deployAddIn
```
とやること。Gradleスクリプトがレポジトリのルートディレクトリの直下にあるxlamファイルを`C:\Users\ユーザ名\AppData\Roaming\Microsoft\AddIns` フォルダにコピーする。
5. けっきょく `C:\Users\ユーザ名\AppData\Roaming\Microsoft\AddIns\kazurayam-vba-lib.xlam` ファイルができる。すでに同名のファイルが存在していたら上書き更新される。
6. レポジトリのルートディレクトリの直下に保存された　kazurayam-vba-lib.xlam　ファイルをgit addしてgit commitしてgit pushしよう。


## AddInを開発するうえでの注意事項

1. `kazurayam-vba-lib.xlsm`ファイルを開いてVBEを開くとプロジェクト名を確認できる。プロジェクト名はユニークな名前にしなければならない。ここでは `KazurayamVbaLib`  とした。新規にワークブックを作ったとし初期状態として`VBAProject`という名前が割り当てられるが、他のブックファイルでも同じプロジェクトが割り当てられる可能性が高い。`VBAProject`が複数あると名前の重複のせいでプロシージャが見つからないというエラーになる。だから `KazurayamVbaLib` というたぶん重複しなさそうなプロジェクト名にして、エラーを回避した。 

## kazurayamがAddInの作り方、利用法を学ぶために参照した情報源

- https://tonari-it.com/excel-vba-class-addin/
- https://tonari-it.com/excel-vba-class-addin-reference/
- https://excel-ubara.com/excelvba4/EXCEL297.html

## VBAソースコードをxlsmファイルからsrcフォルダにどうやってEXPORTしたか

Rubberduckの機能を使った。ctrl+shift+Eでexportできる。