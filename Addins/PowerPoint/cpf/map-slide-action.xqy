xquery version "1.0-ml";
(: Copyright 2009-2010 Mark Logic Corporation

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
(: map-slide-action
::
:: Uses the external variables:
::    $cpf:document-uri: The document being processed
::    $cpf:transition: The transition being executed
::    $cpf:options: 
::       this:map-type (slide|image)
:)
import module namespace cpf = "http://marklogic.com/cpf" at "/MarkLogic/cpf/cpf.xqy";
import module namespace cvt = "http://marklogic.com/cpf/convert" at "/MarkLogic/conversion/convert.xqy";
import module namespace ppt=  "http://marklogic.com/openxml/powerpoint" at "/MarkLogic/openxml/presentation-ml-support.xqy";

declare namespace html = "http://www.w3.org/1999/xhtml";
declare namespace ms="http://schemas.openxmlformats.org/spreadsheetml/2006/main";
declare namespace r="http://schemas.openxmlformats.org/officeDocument/2006/relationships";
declare namespace p = "http://schemas.openxmlformats.org/presentationml/2006/main";
declare namespace this="/MarkLogic/conversion/actions/map-slide-action.xqy";

declare variable $cpf:document-uri as xs:string external;
declare variable $cpf:transition as node() external;
declare variable $cpf:options as node() external;

if (cpf:check-transition($cpf:document-uri,$cpf:transition)) then
try {
  xdmp:trace("Office OpenXML Event", fn:concat("Mapping Slide to Image properties. DOCUMENT-URI: ",$cpf:document-uri )),
  let $map-type := $cpf:options//this:map-type/text()
  let $return := 
   if($map-type eq "image") then
     let $slideprops :=
          if(fn:matches($cpf:document-uri,"Slide\d+\.PNG$")) then
             let $slidetokens := fn:tokenize($cpf:document-uri,"/")
             let $slideimgname := $slidetokens[last()] 
             let $pptx-uri := fn:replace(fn:replace($cpf:document-uri,"_PNG",".pptx"),fn:concat("/",$slideimgname),"")
             let $pptx-dir := fn:replace(fn:replace($cpf:document-uri, $slideimgname,""),"_PNG/","_pptx_parts/")

             let $slidexmluri := ppt:uri-slide-png-to-slide-xml($cpf:document-uri)

             let $slide-idx := fn:replace(fn:replace($slideimgname,"Slide",""),".PNG","")
             return if(fn:empty(fn:doc($slidexmluri))) then
                      () 
                    else 
                      (xdmp:document-set-property($cpf:document-uri, <ppt:pptx>{$pptx-uri}</ppt:pptx>),
                       xdmp:document-set-property($cpf:document-uri, <ppt:pptxdir>{$pptx-dir}</ppt:pptxdir>),
                       xdmp:document-set-property($cpf:document-uri, <ppt:slide>{$slidexmluri}</ppt:slide>),
                       xdmp:document-set-property($cpf:document-uri, <ppt:index>{$slide-idx}</ppt:index>),
                       xdmp:document-set-property($slidexmluri, <ppt:pptx>{$pptx-uri}</ppt:pptx>),
                       xdmp:document-set-property($slidexmluri, <ppt:pptxdir>{$pptx-dir}</ppt:pptxdir>),
                       xdmp:document-set-property($slidexmluri, <ppt:slideimg>{$cpf:document-uri}</ppt:slideimg>),
                       xdmp:document-set-property($slidexmluri, <ppt:index>{$slide-idx}</ppt:index>)
                       )
          else ()
       return $slideprops
   else if($map-type eq "slide") then
      let $doc := fn:doc($cpf:document-uri)
      let $slideprops :=
          if(fn:empty($doc) or fn:not(fn:matches($cpf:document-uri,"slide\d+\.xml$"))) then
             ()
          else
             let $slidetokens := fn:tokenize($cpf:document-uri,"/")
             let $origslidename := $slidetokens[last()]
             let $pptx-uri := fn:replace($cpf:document-uri,"_pptx_parts/ppt/slides/slide\d+\.xml",".pptx")
             let $pptx-dir := fn:replace($cpf:document-uri,"ppt/slides/slide\d+\.xml","")
            
             let $slideimgname := ppt:uri-slide-xml-to-slide-png($cpf:document-uri)

             let $slide-idx := fn:replace(fn:replace($origslidename,"slide",""),".xml","")
             return if(fn:empty(fn:doc($slideimgname))) then 
                      ()
                    else
                      (xdmp:document-set-property($slideimgname, <ppt:pptx>{$pptx-uri}</ppt:pptx>),
                       xdmp:document-set-property($slideimgname, <ppt:pptxdir>{$pptx-dir}</ppt:pptxdir>),
                       xdmp:document-set-property($slideimgname, <ppt:slide>{$cpf:document-uri}</ppt:slide>),
                       xdmp:document-set-property($slideimgname, <ppt:index>{$slide-idx}</ppt:index> ),
                       xdmp:document-set-property($cpf:document-uri, <ppt:pptx>{$pptx-uri}</ppt:pptx>),
                       xdmp:document-set-property($cpf:document-uri, <ppt:pptxdir>{$pptx-dir}</ppt:pptxdir>),
                       xdmp:document-set-property($cpf:document-uri, <ppt:slideimg>{$slideimgname}</ppt:slideimg>),
                       xdmp:document-set-property($cpf:document-uri, <ppt:index>{$slide-idx}</ppt:index> )
                      )
      return $slideprops
   else ()
  return ($return,
  cpf:success( $cpf:document-uri, $cpf:transition, ())) 
  
}catch ($e) {
   cpf:failure( $cpf:document-uri, $cpf:transition, $e, () )
}
else ()
