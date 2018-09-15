# kamiWord: Officeテンプレートに値を入れて新しいドキュメントを生成する

## これは何？

これは，docx形式のWordテンプレートまたはxlsx形式のExcelテンプレートに値を入れて新たなドキュメントを生成するためのTclライブラリです．

## 何がうれしいの？
できることはWordの差し込み印刷と似ていますが，VBA以外の言語から操作できること，複数のファイルを生成できることなどが違います．

## 必要なもの
- Tcl 8.5以上
- zip, unzip コマンド（Windows標準のもの，またはそれと同じ呼び出し形式のもの）
- fileutilパッケージ（tcllibに附属）
- cmdlineパッケージ（tcllibに附属）

## インストール
Tclのライブラリサーチパス上に適当なディレクトリを掘り，kamiWord.tcl とpkgIndex.tcl をコピーします．

## 使用法

### ふつうの使用法
1. まずテンプレートを用意します．任意のdocx形式Word文書またはxlsx形式Excelワークシートの中に，プレースホルダーとして !!!キー!!! という文字列を挿入します（「キー」の部分は適当な文字列）．
2. Tclスクリプトの中で，キーとそこに流し込む値の対応を定義します．

```
set values [::kamiWord::replacelist]
::kamiWord::setvalue values キー 適当な値
```

3. 文書を開き，次のようにして新しい文書を生成します．

```
set word [::kamiWord::open_document テンプレートファイル名]
::kamiWord::replace $word $values
::kamiWord::finish $word 出力ファイル名
```

### ネ申への道
いわゆる「ネ申Excel」では，1つの値を複数のセルに1文字ずつ入れたいことがあります．このような場合，テンプレートに!!!キー0!!!, !!!キー1!!!，のように0から始まる添字を付けたプレースホルダーを挿入し，値をセットするときに

```
::kamiWord::setvalue -explode values キー １２３４５６
```

のようにすることで，セットする値を1文字ずつ分解し，キーに添え字を付けてセットします．1文字ずつではなく，何かの文字で区切りたい場合は

```
::kamiWord::setvalue -explode -sepchar - キー 012-345-6789
```

のようにして区切り記号を指定することができます．

## どうやっているのか
WordやExcelなどの"docx"や"xlsx"などの形式，いわゆるOOXML形式ドキュメントは，XMLのドキュメントと画像などのメディアファイルをフォルダ内に配置したものを単にzip圧縮したファイルです．Word文書であるdocx形式の場合，ドキュメント内の文字列に対応しているのは word/document.xml です．そこで，word/document.xml内のプレースホルダーの文字列を単に文字列置換して，zipで固め直しているだけです．

現在，改行がうまく表現できていません．改行文字は&amp;#10;で表現されていますが，ドキュメントの文字列だけでなくフォーマットも操作しなければならないため，単に文字列の一部を&amp;#10;にしただけだと改行になってくれません．（対応策の提案募集中）

## 注意点
`::kamiWord::open_document` で文書を開くと，その文書の中身を一時ファイル用フォルダ（典型的にはユーザフォルダ直下のAppData/Local/Temp）に展開します．`::kamiWord::finish`でその内容を削除しますが，`::kamiWord::finish`を実行せずにプログラムを終了すると一時ファイルがそのまま残ります．一時ファイルは"プロセス番号.通番"という名前になっているので，同じプロセス番号で再びプログラムが動いた時にファイルが展開できずにエラーになることがあります．

## 参考文献
Ito, Akinori. "Demonstration Experiment of Data Hiding into OOXML Document for Suppression of Plagiarism." Advances in Intelligent Information Hiding and Multimedia Signal Processing. Springer, Cham, 2017. 3-10.

## 作者
伊藤彰則 aito@fw.ipsj.or.jp
