xquery version "1.0-ml";

let $docname := xdmp:get-request-field("url")
let $title := xdmp:get-request-field("title")

let $package := fn:doc($docname)

    let $filename := $title 
    let $disposition := concat("attachment; filename=""",$filename,"""")
    let $x := xdmp:add-response-header("Content-Disposition", $disposition)
    let $x:= if(fn:contains($filename,"docx")) then 
              xdmp:set-response-content-type("application/vnd.openxmlformats-officedocument.wordprocessingml.document")
         else if(fn:contains($filename,"pptx")) then
              xdmp:set-response-content-type("application/vnd.openxmlformats-officedocument.presentationml.presentation") 
         else if(fn:contains($filename,"xlsx")) then
              xdmp:set-response-content-type("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
         else if(fn:contains($filename,"docm")) then
              xdmp:set-response-content-type("application/vnd.ms-word.document.macroEnabled.12")
         else if(fn:contains($filename,"dotm")) then
              xdmp:set-response-content-type("application/vnd.ms-word.template.macroEnabled.12")
         else if(fn:contains($filename,"dotx")) then
              xdmp:set-response-content-type("application/vnd.openxmlformats-officedocument.wordprocessingml.template")
         else if(fn:contains($filename,"ppsm")) then
              xdmp:set-response-content-type("application/vnd.ms-powerpoint.slideshow.macroEnabled.12")
         else if(fn:contains($filename,"ppsx")) then
              xdmp:set-response-content-type("application/vnd.openxmlformats-officedocument.presentationml.slideshow")
         else if(fn:contains($filename,"pptm")) then
              xdmp:set-response-content-type("application/vnd.ms-powerpoint.presentation.macroEnabled.12")
         else if(fn:contains($filename,"xlsb")) then
              xdmp:set-response-content-type("application/vnd.ms-excel.sheet.binary.macroEnabled.12")
         else if(fn:contains($filename,"xlsm")) then
              xdmp:set-response-content-type("application/vnd.ms-excel.sheet.macroEnabled.12")
         else  (: .xps  :)
              xdmp:set-response-content-type("application/vnd.ms-xpsdocument")
    return
      $package

