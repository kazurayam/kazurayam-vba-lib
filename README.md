# kazurayam-vba-lib

A Library of VBA procedures created by kazurayam for his own use

kazurayamが自分が使うために作ったVBAアドインのライブラリ。

## AddInにかんする技術情報

- https://tonari-it.com/excel-vba-class-addin/
- https://tonari-it.com/excel-vba-class-addin-reference/
- https://excel-ubara.com/excelvba4/EXCEL297.html


## AddInを開発するうえでの注意事項

1. `kazurayam-vba-lib.xlsm`ファイルを開いてVBEを開くとプロジェクト名を確認できる。プロジェクト名はユニークな名前にしなければならない。ここでは `KazurayamVbaLib`  とした
2. デフォルトとして`VBAProject`という名前が割り当てられるが、他のブックファイルでも同じプロジェクトが割り当てられるかもしれない。`VBAProject`が複数あると名前の重複のせいでプロシージャが見つからないというエラーになる。

## AddInを配置する手順

1. `kazurayam-vba-lib.xlsm` ファイルをExcelで開く。
2. `名前を付けて保存`を選択する。出力するファイルの形式として `xlam` を選択する。すると保存先フォルダとして `C:\Users\ユーザ名\AppData\Roaming\Microsoft\AddIns` フォルダが自動的に選択されるので、そのまま保存する。
3. けっきょく `C:\Users\ユーザ名\AppData\Roaming\Microsoft\AddIns\kazurayam-vba-lib.xlam` ファイルができる。すでに同名のファイルが存在していたら上書き更新される。

## 別のエクセルワークブックがkazurayam-vba-lib.xlamのAddInを利用する手順

1. エクセルワークブックを開く。
2. VBEを開く。
3. VBEのツールバーで　ツール＞参照設定　を選ぶ。参照設定のダイアログが表示される。
4. 参照設定のダイアログの右側にあるボタンのうち　参照　ボタンをクリックする。
5. `C:\Users\ユーザ名\AppData\Roaming\Microsoft\AddIns\kazurayam-vba-lib.xlam`　を選択する。OKボタンを押下する。