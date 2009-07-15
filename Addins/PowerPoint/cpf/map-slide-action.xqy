xquery version "1.0-ml";
(: Copyright 2002-2009 Mark Logic Corporation.  All Rights Reserved. :)

(: map-sharedstrings
::
:: Uses the external variables:
::    $cpf:document-uri: The document being processed
::    $cpf:transition: The transition being executed
:)
import module namespace cpf = "http://marklogic.com/cpf" at "/MarkLogic/cpf/cpf.xqy";
import module namespace cvt = "http://marklogic.com/cpf/convert" at "/MarkLogic/conversion/convert.xqy";
declare namespace excel = "http://marklogic.com/openxml/excel";
import module "http://marklogic.com/openxml/excel" at "/MarkLogic/openxml/spreadsheet-ml-support.xqy";
declare namespace html = "http://www.w3.org/1999/xhtml";
declare namespace ms="http://schemas.openxmlformats.org/spreadsheetml/2006/main";
declare namespace r="http://schemas.openxmlformats.org/officeDocument/2006/relationships";
declare namespace p = "http://schemas.openxmlformats.org/presentationml/2006/main";

declare variable $cpf:document-uri as xs:string external;
declare variable $cpf:transition as node() external;

if (cpf:check-transition($cpf:document-uri,$cpf:transition)) then
try {
  xdmp:trace("Office OpenXML Event", fn:concat("Mapping Slide to Image properties. DOCUMENT-URI: ",$cpf:document-uri )),
  let $doc := fn:doc($cpf:document-uri)
  let $ret :=  if(fn:empty($doc) or fn:contains($cpf:document-uri,"slideLayout")) 
               then ()
               else
               (: let $cpfuri := "/Aven_MarkLogicUserConference2009Exceling_pptx_parts/ppt/slides/slide1.xml" :)

               let $pptxuri := fn:replace($cpf:document-uri,"_pptx_parts/ppt/slides/slide\d+\.xml",".pptx")
               let $slidetokens := fn:tokenize($cpf:document-uri,"/")
               let $origslidename := $slidetokens[last()]
               let $slidedir := fn:replace(fn:replace(fn:replace($cpf:document-uri,$origslidename,""),"/ppt/slides/",""),"_pptx_parts","_PNG/")
               let $slidename := fn:concat($slidedir,fn:replace(fn:replace($origslidename,"slide","Slide"),".xml",".PNG"))

(:check fn:doc($slidename), if this is empty, need to not set doc props :)

               let $slideidx := fn:replace(fn:replace($origslidename,"slide",""),".xml","")
               return (xdmp:document-set-properties($slidename ,(<pptx>{$pptxuri}</pptx>,
                                                                 <slide>{$cpf:document-uri}</slide>,
                                                                 <index>{$slideidx}</index>)),
                       xdmp:document-set-properties($cpf:document-uri, (<slideimg>{$slidename}</slideimg>))
                       )
(:
               let $imgprops := xdmp:document-set-properties($slidename ,(<pptx>{$pptxuri}</pptx>,
                                                                          <slide>{$cpf:document-uri}</slide>,
                                                                          <index>{$slideidx}</index>))

               let $slideprops := xdmp:document-set-properties($cpfuri, (<slideimg>{$slidename}</slideimg>))

               return   ($cpfuri,$pptxuri ,$origslidename, $slidename, $slidedir, $slideidx)
:)
(:
               let $uri := fn:concat("/foo",$cpf:document-uri)
               return xdmp:document-insert($uri, <fubar>some junk</fubar>)
:)
(:
               let $uri:= $cpf:document-uri
               let $path := fn:tokenize($uri,"/")
               let $dirs := $path[last()]
               let $count := fn:count($path)-1

               let $sharedstringdir := for $i at $d in $path
                                       let $string :=  if($i eq "") then () else $i
                                       where $d lt $count
                                       return $string

               let $ssuri := fn:concat("/",fn:string-join($sharedstringdir,"/"),"/sharedStrings.xml")

               let $sheet:= fn:doc($uri)/node()
               let $shared-strings :=  fn:doc($ssuri)/node()
               let $newsheet := if(fn:empty($shared-strings)) then $sheet else excel:map-shared-strings($sheet, $shared-strings)
  
               return xdmp:document-insert($cpf:document-uri,$newsheet)
:)
 
  return ($ret,
  cpf:success( $cpf:document-uri, $cpf:transition, ())) 
  
}catch ($e) {
   cpf:failure( $cpf:document-uri, $cpf:transition, $e, () )
}
else ()
