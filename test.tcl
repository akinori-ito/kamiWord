lappend auto_path .
package require kamiWord

set template template.docx
set template2 template.xlsx

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

set values [dict create]
dict set values {!!!x!!!} その１
dict set values {!!!y!!!} その２
set excel [::kamiWord::open_docx $template2]
::kamiWord::replace $excel $values
::kamiWord::finish $excel replaced.xlsx