lappend auto_path .
package require kamiWord

set template template.docx
set template2 template.xlsx

set values [::kamiWord::replacelist]
::kamiWord::setvalue values Name ほげ太郎
::kamiWord::setvalue values MF 男
::kamiWord::setvalue values PH 012-345-6789
::kamiWord::setvalue values Address {東京都千代田区丸の内1-1}
::kamiWord::setvalue values Tech {Tclでプログラムを書く}
::kamiWord::setvalue values Veg ○
::kamiWord::setvalue values Cat ○

set word [::kamiWord::open_document $template]
::kamiWord::replace $word $values
::kamiWord::finish $word replaced.docx

set values [::kamiWord::replacelist]
::kamiWord::setvalue values Name ほげ太郎
::kamiWord::setvalue values Age 32
::kamiWord::setvalue -explode values 〒 1000000
::kamiWord::setvalue values Address {東京都千代田区丸の内1-1}
::kamiWord::setvalue -explode -sepchar - values Tel 012-345-6789
set excel [::kamiWord::open_document $template2]
::kamiWord::replace $excel $values
::kamiWord::finish $excel replaced.xlsx