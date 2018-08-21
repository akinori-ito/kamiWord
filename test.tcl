lappend auto_path .
package require kamiWord

set template template.docx
set template2 template.xlsx

set values [dict create]
dict set values {!!!Name!!!} Ù°¾Y
dict set values {!!!MF!!!} j
dict set values {!!!PH!!!} 012-345-6789
dict set values {!!!Address!!!} {sçãcæÛÌà1-1}
dict set values {!!!Tech!!!} {TclÅvOð­}
dict set values {!!!Veg!!!} 
dict set values {!!!Cat!!!} 

set word [::kamiWord::open_docx $template]
::kamiWord::replace $word $values
::kamiWord::finish $word replaced.docx

set values [dict create]
dict set values {!!!x!!!} »ÌP
dict set values {!!!y!!!} »ÌQ
set excel [::kamiWord::open_docx $template2]
::kamiWord::replace $excel $values
::kamiWord::finish $excel replaced.xlsx