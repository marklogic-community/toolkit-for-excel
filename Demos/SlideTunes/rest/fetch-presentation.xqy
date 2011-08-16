xquery version "1.0-ml";

import module namespace rest = "http://marklogic.com/appservices/rest"
        at "/MarkLogic/appservices/utils/rest.xqy";


import module namespace requests =   "http://marklogic.com/appservices/requests" at "requests.xqy";

let $request := $requests:options/rest:request[@endpoint = "/fetch-presentation.xqy"][1]

let $map  := rest:process-request($request) 

let $deck := map:get($map, "deck")
let $filename := fn:tokenize($deck,"/")[last()]
let $fullpath := fn:substring-after($deck,"/office/presentations")

let $x-package := fn:doc($fullpath)
let $disposition := concat("attachment; filename=""",$filename,"""")
let $y := xdmp:add-response-header("Content-Disposition", $disposition)
let $y := if(fn:contains($filename,"docx")) then 
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
    $x-package
