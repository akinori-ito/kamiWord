lappend auto_path .
package require kamiWord

set template template.docx

set values [dict create]
dict set values {!!!Name!!!} ほげ太郎
dict set values {!!!MF!!!} 男
dict set values {!!!PH!!!} 012-345-6789
dict set values {!!!Address!!!} {東京都千代田区丸の内1-1}
dict set values {!!!Tech!!!} {Tclでプログラムを書く}
dict set values {!!!Veg!!!} ○
dict set values {!!!Cat!!!} ○

set word [::kamiWord::open_docx $template]
::kamiWord::replace $word $values
::kamiWord::finish $word replaced.docx
