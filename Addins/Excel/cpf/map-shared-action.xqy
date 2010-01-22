xquery version "1.0-ml";
(: Copyright 2002-2010 Mark Logic Corporation.  All Rights Reserved. :)

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

declare variable $cpf:document-uri as xs:string external;
declare variable $cpf:transition as node() external;

if (cpf:check-transition($cpf:document-uri,$cpf:transition)) then
try {
  xdmp:trace("Office OpenXML Event", fn:concat("Mapping SharedStrings. DOCUMENT-URI: ",$cpf:document-uri )),
  let $doc := fn:doc($cpf:document-uri)
  let $ret :=  if(fn:empty($doc)) then ()
else

  let $uri:= $cpf:document-uri (:  "/Default_xlsx_parts/xl/worksheets/sheet1.xml" :)
  let $path := fn:tokenize($uri,"/")
  let $dirs := $path[last()]
  let $count := fn:count($path)-1

  let $sharedstringdir := for $i at $d in $path
                          let $string :=  if($i eq "") then () else $i
                          where $d lt $count
                          return $string

  let $ssuri := fn:concat("/",fn:string-join($sharedstringdir,"/"),"/sharedStrings.xml")

  let $sheet:= fn:doc($uri)/node()
  (: let $shared-strings :=  fn:data(fn:doc($ssuri)//ms:t) :)
  let $shared-strings :=  fn:doc($ssuri)/node()
  let $newsheet := if(fn:empty($shared-strings)) then $sheet else excel:map-shared-strings($sheet, $shared-strings)
  
  return xdmp:document-insert($cpf:document-uri,$newsheet)
 
        (: worked: xdmp:document-insert($cpf:document-uri,$doc) :)
        (: reference from merge -runs
        let $mapped-doc := ooxml:runs-merge($doc/element())
        return xdmp:document-insert($cpf:document-uri,$mapped-doc)
         :)
  return ($ret,
  cpf:success( $cpf:document-uri, $cpf:transition, ())) 
  
}catch ($e) {
   cpf:failure( $cpf:document-uri, $cpf:transition, $e, () )
}
else ()
