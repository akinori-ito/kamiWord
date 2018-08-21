package provide kamiWord 1.0
package require fileutil

namespace eval kamiWord {
  variable tempdir [::fileutil::tempdir]
  variable seq 0
  variable holderpattern {!!![^!]*!!!}
  
  proc replace {zdir vals} {
    variable holderpattern
    set curdir [pwd]
    cd $zdir
    set f [::open word/document.xml r]
    fconfigure $f -encoding utf-8
    set content [read $f]
    close $f
    foreach k [dict keys $vals] {
      set value [dict get $vals $k]
      regsub -all "&" $value {\\\&amp;} value
      regsub -all "<" $value {\\\&lt;} value
      regsub -all ">" $value {\\\&gt;} value
      regsub -all $k $content $value content
    }
    regsub -all $holderpattern $content {} content
    set f [open word/document.xml w]
    fconfigure $f -encoding utf-8
    puts -nonewline $f $content
    close $f
    cd $curdir
 }
 
 proc open_docx {docxfile} {
    variable tempdir
    variable seq
    set zipdir $tempdir/[pid].$seq
    incr seq
    exec unzip $docxfile -d $zipdir
    return $zipdir
 }
 
 proc finish {zdir docxfile} {
    set destfile [file normalize $docxfile]
    if {[file exists $destfile]} {
      file delete $destfile
    }
    set curdir [pwd]
    cd $zdir
    exec zip -r $destfile customXml docProps word _rels {[Content_Types].xml}
    cd $curdir
    file delete -force $zdir
 }
   
}
