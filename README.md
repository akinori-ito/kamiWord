# kamiWord: wordテンプレートに値を入れて新しいドキュメントを生成する

## これは何？

これは，docx形式のWordテンプレートに値を入れて新たなWordドキュメントを生成するためのTclライブラリです．

## 何がうれしいの？
できることはWordの差し込み印刷と似ていますが，VBA以外の言語から操作できること，複数のファイルを生成できることなどが違います．

## 必要なもの
- Tcl 8.5以上
- zip, unzip コマンド（Windows標準のもの，またはそれと同じ呼び出し形式のもの）
- fileutilパッケージ（tcllibに附属）

## インストール
Tclのライブラリサーチパス上に適当なディレクトリを掘り，kamiWord.tcl とpkgIndex.tcl をコピーします．

## 使用法

1. まずテンプレートを用意します．任意のdocx形式Word文書の中に，プレースホルダーとして !!!キー!!! という文字列を挿入します（「キー」の部分は適当な文字列）．
2. Tclスクリプトの中で，キーとそこに流し込む値の対応を定義します．対応にはdictを使います．

```
set values [dict create]
dict set values {!!!キー!!!} 適当な値
```

3. Word文書を開き，次のようにして新しい文書を生成します．

```
set word [::kamiWord::open_word テンプレートファイル名]
::kamiWord::replace $word $values
::kamiWord::finish $word 出力ファイル名
```

## 作者
伊藤彰則 aito@fw.ipsj.or.jp
