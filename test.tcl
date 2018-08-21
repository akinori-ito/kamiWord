lappend auto_path .
package require kamiWord

set template template.docx

set values [dict create]
dict set values {!!!Name!!!} �ق����Y
dict set values {!!!MF!!!} �j
dict set values {!!!PH!!!} 012-345-6789
dict set values {!!!Address!!!} {�����s���c��ۂ̓�1-1}
dict set values {!!!Tech!!!} {Tcl�Ńv���O����������}
dict set values {!!!Veg!!!} ��
dict set values {!!!Cat!!!} ��

set word [::kamiWord::open_docx $template]
::kamiWord::replace $word $values
::kamiWord::finish $word replaced.docx
