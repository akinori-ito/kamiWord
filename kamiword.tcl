package provide kamiWord 1.0
package require fileutil
package require cmdline

namespace eval kamiWord {
  variable tempdir [::fileutil::tempdir]
  variable seq 0
  variable holderpattern {!!![^!]*!!!}
  
  proc replacelist {} {
     return [dict create]
  }
  
  proc setvalue args {
    ::set arglist {{explode "Explode value"} {sepchar.arg {} "character to separate value"}}
    ::set usage {Usage: ::kamiWord::set [-explode] [-sepchar sepchar] replacelistvar key value}
    ::set myarg $args
    array set params [::cmdline::getoptions myarg $arglist $usage]
    if {[llength $myarg] != 3} {
     error "$usage"
    }
    upvar [lindex $myarg 0] replacelist
    ::set key [lindex $myarg 1]
    ::set value [lindex $myarg 2]
    if {$params(explode)} {
     set explodearg [split $value $params(sepchar)]
     for {set i 0} {$i < [llength $explodearg]} {incr i} {
      dict set replacelist "!!!$key$i!!!" [lindex $explodearg $i]
     }
    } else {
     dict set replacelist "!!!$key!!!" $value
    }
  }
  
  proc replace {zdir vals} {
    variable holderpattern
    set curdir [pwd]
    cd [lindex $zdir 0]
    set f [::open [lindex $zdir 1] r]
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
    set f [open [lindex $zdir 1] w]
    fconfigure $f -encoding utf-8
    puts -nonewline $f $content
    close $f
    cd $curdir
 }
 
 proc open_document {docxfile} {
    variable tempdir
    variable seq
    switch [file extension $docxfile] {
     .docx  {
      set replacefile word/document.xml
      set contentlist {customXml docProps word _rels {[Content_Types].xml}}
      }
     .xlsx  {
      set replacefile xl/sharedStrings.xml
      set contentlist {_rels docProps xl {[Content_Types].xml}}
      }
     default { error "Bad file extension [file extension $docxfile]" }
    }
    set zipdir $tempdir/[pid].$seq
    incr seq
    exec unzip $docxfile -d $zipdir
    return [list $zipdir $replacefile $contentlist]
 }
 
 proc finish {zdir docxfile} {
    set destfile [file normalize $docxfile]
    if {[file exists $destfile]} {
      file delete $destfile
    }
    set curdir [pwd]
    cd [lindex $zdir 0]
    eval "exec zip -r $destfile [lindex $zdir 2]"
    cd $curdir
    file delete -force $zdir
 }
   
}
