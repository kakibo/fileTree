# fileTree
## 使い方
node fileTree [options] <Target Directory> [output excel path & name]

-p, --path [url|pc|pcfull]
    パスの表示方法
    url　   unix形式　デフォルト
    pc  　　windows形式の相対パス
    pcfull  windows形式の絶対パス
-f, --filter ["ext|ext"]
    拡張子フィルタ ※ダブルクォーテーションで囲み、半角「|」区切りで拡張子を指定してください。
-s, --search ["正規表現"]
    簡易検索機能。該当文字列があるファイルにメモが追記されます。
    検索文字列を指定。正規表現が利用できます。※ダブルクォーテーションで囲んでください。
-T, --test
    デバッグモード(詳細ログを出力）

example:
HTMLだけを検索し、サイトマップを作る
node fileTree -f "html" c:\path\to\folder
