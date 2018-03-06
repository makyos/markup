# markup

エクセルシートの選択中セル内容を、テーブルタグにしてクリップボードに出力する、 "Excel Addin" です。

## セットアップ

### インストール

MarkUp.xla を、下記のディレクトリに入れる。

%HOMEPATH%\AppData\Roaming\Microsoft\AddIns

### アドイン有効化

Excel ＞ 開発タブ ＞ アドイン で、以下のファイルを「参照」設定する。

MarkUp.xla

### 参照設定

Excel ＞ 開発タブ ＞ Visual Basic > 参照設定 で、以下のファイルを「参照」設定する。

c:\Windows\System32\FM20.DLL

## 使い方

テーブルタグにしたいセル範囲を選択し、コンテキストメニューから "CopyHTML" を実行すると、結果がクリップボードに保存される。


以上
