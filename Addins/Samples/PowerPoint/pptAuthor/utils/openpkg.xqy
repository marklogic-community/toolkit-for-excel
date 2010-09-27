xquery version "1.0-ml";
(:
Copyright 2008-2010 Mark Logic Corporation

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.

:)

declare namespace w="http://schemas.openxmlformats.org/wordprocessingml/2006/main";

let $doc-uri := xdmp:url-decode(xdmp:get-request-field('uri'))
let $docx-uri := if(fn:contains($doc-uri,"_parts/word/document.xml")) then 
                    fn:concat(fn:substring-before($doc-uri,"_docx_parts"),".docx")
                 else if (fn:contains($doc-uri,"ppt/slides")) then
                    fn:concat(fn:substring-before($doc-uri,"_pptx_parts"),".pptx")
                 else
                    $doc-uri 

let $filename :=  fn:tokenize($docx-uri,"/")[last()]

let $disposition := fn:concat("attachment; filename=""",$filename,"""")
let $x := xdmp:add-response-header("Content-Disposition", $disposition)
let $x:= if(fn:contains($filename, ".xml")) then
              xdmp:set-response-content-type("application/vnd.openxmlformats-officedocument.wordprocessingml.document")
         else if(fn:contains($filename,"docx")) then 
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
return	fn:doc($docx-uri)
	
