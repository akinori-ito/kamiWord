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
set word [::kamiWord::open_docx テンプレートファイル名]
::kamiWord::replace $word $values
::kamiWord::finish $word 出力ファイル名
```

## どうやっているのか
WordやExcelなどの"docx"や"xlsx"などの形式，いわゆるOOXML形式ドキュメントは，XMLのドキュメントと画像などのメディアファイルをフォルダ内に配置したものを単にzip圧縮したファイルです．Word文書であるdocx形式の場合，ドキュメント内の文字列に対応しているのは word/document.xml です．そこで，word/document.xml内のプレースホルダーの文字列を単に文字列置換して，zipで固め直しているだけです．

現在，改行がうまく表現できていません．改行文字は&#10;で表現されていますが，ドキュメントの文字列だけでなくフォーマットも操作しなければならないため，単に文字列の一部を&#10;にしただけだと改行になってくれません．（対応策の提案募集中）

## 参考文献
Ito, Akinori. "Demonstration Experiment of Data Hiding into OOXML Document for Suppression of Plagiarism." Advances in Intelligent Information Hiding and Multimedia Signal Processing. Springer, Cham, 2017. 3-10.

## 作者
伊藤彰則 aito@fw.ipsj.or.jp
