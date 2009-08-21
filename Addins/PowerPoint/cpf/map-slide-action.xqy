xquery version "1.0-ml";
(: Copyright 2002-2009 Mark Logic Corporation.  All Rights Reserved. :)

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
declare namespace excel = "http://marklogic.com/openxml/excel";
import module "http://marklogic.com/openxml/excel" at "/MarkLogic/openxml/spreadsheet-ml-support.xqy";
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
      let $image := 
      if(fn:matches($cpf:document-uri,"Slide\d+\.PNG$")) then
        let $slidetokens := fn:tokenize($cpf:document-uri,"/")
        let $slideimgname := $slidetokens[last()]
        let $slidexmlname := fn:replace(fn:replace($slideimgname,"S","s"),"PNG","xml")
        let $slidedir := 
fn:replace(fn:replace($cpf:document-uri,$slideimgname,""),"_PNG","_pptx_parts")
        let $slidexmluri := fn:concat($slidedir,"ppt/slides/",$slidexmlname) 
        let $props := 
             if(fn:empty(fn:doc($slidexmluri))) then
               () 
             else 
               let $pptx-uri := fn:replace(fn:replace($cpf:document-uri,"_PNG",".pptx"),fn:concat("/",$slideimgname),"")
               let $slide-idx := fn:replace(fn:replace($slideimgname,"Slide",""),".PNG","")
               return (xdmp:document-set-properties($cpf:document-uri,(<pptx>{$pptx-uri}</pptx>,
                                                                       <slide>{$slidexmluri}</slide>,
                                                                        <index>{$slide-idx}</index>)),
                       xdmp:document-set-properties($slidexmluri, (<slideimg>{$cpf:document-uri}</slideimg>)))
        return $props (: ($cpf:document-uri,$props, $slideimgname, $slidexmlname,$slidexmluri ,$slidedir)  :)
      else ()
      return $image
   else if($map-type eq "slide") then
      let $doc := fn:doc($cpf:document-uri)
      let $slide :=
          if(fn:empty($doc) or fn:not(fn:matches($cpf:document-uri,"slide\d+\.xml$"))) then
             ()
          else
             let $pptxuri := fn:replace($cpf:document-uri,"_pptx_parts/ppt/slides/slide\d+\.xml",".pptx")
             let $slidetokens := fn:tokenize($cpf:document-uri,"/")
             let $origslidename := $slidetokens[last()]
             let $slidedir := fn:replace(fn:replace(fn:replace($cpf:document-uri,$origslidename,""),"/ppt/slides/",""),"_pptx_parts","_PNG/")
             let $slideimgname := fn:concat($slidedir,fn:replace(fn:replace($origslidename,"slide","Slide"),".xml",".PNG"))
             let $slideidx := fn:replace(fn:replace($origslidename,"slide",""),".xml","")
             return if(fn:empty(fn:doc($slideimgname))) then 
                      ()
                    else
                      (xdmp:document-set-properties($slideimgname ,(<pptx>{$pptxuri}</pptx>,
                                                                 <slide>{$cpf:document-uri}</slide>,
                                                                 <index>{$slideidx}</index>)),
                       xdmp:document-set-properties($cpf:document-uri, (<slideimg>{$slideimgname}</slideimg>))
                       )
      return $slide
               
   else ()
  return ($return,
  cpf:success( $cpf:document-uri, $cpf:transition, ())) 
  
}catch ($e) {
   cpf:failure( $cpf:document-uri, $cpf:transition, $e, () )
}
else ()
