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
declare namespace this="/MarkLogic/conversion/actions/pptx-set-tags-action.xqy";

declare variable $cpf:document-uri as xs:string external;
declare variable $cpf:transition as node() external;
declare variable $cpf:options as node() external;

if (cpf:check-transition($cpf:document-uri,$cpf:transition)) then
try {
  xdmp:trace("Office OpenXML Event", fn:concat("Setting document properties with Tag information. DOCUMENT-URI: ",$cpf:document-uri )),
  let $tag-type := $cpf:options//this:tag-type/text()
  let $docname := $cpf:document-uri
  let $doc:= fn:doc($docname)
  let $return := 
       if(fn:empty($doc//p:tags)) then ()
       else
         let $doctokens := fn:tokenize($docname,"/")
         let $origdocname := $doctokens[last()]

         let $docrels := 
             if(fn:ends-with($docname,"presentation.xml")) then
                fn:concat("_rels/","presentation.xml.rels")
             else
                let $slide-idx := fn:replace(fn:replace($origdocname,"slide",""),".xml","")
                return fn:concat("_rels/slide",$slide-idx,".xml.rels")

         let $relsuri := fn:replace(fn:string-join($doctokens,"/"),$origdocname,$docrels)

         let $reldoc := fn:doc($relsuri)/node() 

         let $props := for $tag in $doc//p:tags
                       return ppt:create-tag-properties($docname, $tag, $reldoc)

         let $finalprops := for $p in $props
                            return if(fn:local-name($p) eq "slidetags" or fn:local-name($p) eq "presentationtags") then $p else ()

         let $justshapes := for $elem at $d in $props
                    return if(fn:local-name($elem) eq "shapetags") then
                        $elem
                    else ()
                    
          
         let $finalshapes := if(fn:not(fn:empty($justshapes))) then
                                <ppt:shapetags>
                                   {for $shp in $justshapes
                                    return <ppt:shape>{$shp/ppt:tag }</ppt:shape>
                                   }
                                </ppt:shapetags>
                              else ()
 
         let $final := ($finalprops, $finalshapes)

         return  if(fn:ends-with($docname,"presentation.xml")) then
                    let $preso-props :=  xdmp:document-set-property($docname, $final) 
                    let $slide-dir := fn:concat(fn:substring-before($docname,"presentation.xml"),"slides/")
                    let $slides := cts:search(fn:collection(),cts:directory-query($slide-dir,"1"))
                    let $return := for $s in $slides
                                   let $uri := xdmp:node-uri($s)
                                   return xdmp:document-set-property($uri,$final)
                    return $return 
                 else
                     xdmp:document-set-property($docname, $final) 

  return ($return,
  cpf:success( $cpf:document-uri, $cpf:transition, ())) 
  
}catch ($e) {
   cpf:failure( $cpf:document-uri, $cpf:transition, $e, () )
}
else ()
