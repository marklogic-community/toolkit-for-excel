xquery version "1.0-ml";
(:
Copyright 2009-2010 Mark Logic Corporation

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
import module namespace ppt=  "http://marklogic.com/openxml/powerpoint" at "/MarkLogic/openxml/presentation-ml-support.xqy";

declare namespace pptadd="http://marklogic.com/openxml/pptaddin";
declare namespace html = "http://www.w3.org/1999/xhtml";
declare namespace a="http://schemas.openxmlformats.org/drawingml/2006/main";
declare namespace r="http://schemas.openxmlformats.org/officeDocument/2006/relationships";
declare namespace p="http://schemas.openxmlformats.org/presentationml/2006/main";
declare namespace dc = "http://purl.org/dc/elements/1.1/";
declare namespace cp = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
declare namespace dcterms="http://purl.org/dc/terms/";

declare variable $pptadd:bsv as xs:string external;
declare variable $pptadd:searchtype as xs:string external;

let $searchparam := $pptadd:bsv
let $searchtype :=  $pptadd:searchtype

let $return := 
if($searchtype eq "slide") then
     let $slides := cts:search(//p:sld, cts:word-query($searchparam))

     let $slideuris := for $s in $slides 
                       let $orig-uri := xdmp:node-uri($s)
                       return xdmp:document-properties($orig-uri)/prop:properties/ppt:slideimg/text()
     let $disp-slides := 
         for $pic at $d in $slideuris
         let $src := fn:concat("download-support.xqy?uid=",$pic)

         let $prop := xdmp:document-properties($pic)
         let $pptx := $prop/prop:properties/ppt:pptx/text()
         let $slide := $prop/prop:properties/ppt:slide/text()
         let $index := $prop/prop:properties/ppt:index/text()

         let $imganchor := fn:concat("#num",$d)
         let $imgnum := fn:concat("num",$d)
         return
          (
              <div>
              <li>
                <table>
                  <tr>
                    <td><a name="{$imgnum}" href="{$imganchor}" onclick="insertSlide('{$pptx}','{$index}','{$d}')">
                          <img src="{$src}" class="resize"></img>
                        </a>
                    </td>
                    <td style="vertical-align: top;" >
                          <input type="checkbox" id="{fn:concat("retain",$d)}" name="format"/>retain format
                    </td>
                 </tr>
                </table>
              </li>
              <br/>
              </div>,<br/> 
          )
     
     return <div><ul class="thumb">{$disp-slides}</ul></div>
else if($searchtype eq "image") then
     let $pics := cts:search(fn:collection(), cts:properties-query($searchparam)) 
                         (: cts:uri-match(fn:concat("/",$searchparam,"*.jpg")) :)
     for $pic at $d in $pics
       let $uri := xdmp:node-uri($pic)
       let $src := fn:concat("download-support.xqy?uid=",$uri)
       let $imganchor := fn:concat("#num",$d)
       let $imgnum := fn:concat("num",$d) 

       return 
         (<a name="{$imgnum}" href="{$imganchor}" onclick="insertImage('{$uri (:$imageuri:)}')">
          <img src="{$src}"></img>
          </a>,<br/>,<br/>)
else
let $slides := cts:search(//p:sld, cts:word-query($searchparam))
let $docuris := for $s in $slides 
                let $orig-uri := xdmp:node-uri($s)
                return xdmp:document-properties($orig-uri)/prop:properties/ppt:pptx/text()
let $finaldocs := for $doc in  fn:distinct-values($docuris)
                  let $docfolder := fn:replace($doc,".pptx","_pptx")
                  let $props := fn:concat($docfolder,"_parts/docProps/core.xml")
                  let $propsdoc := fn:doc($props)
                            let $lastmodby := if(fn:empty($propsdoc//cp:lastModifiedBy//text())) then () 
                                              else fn:concat("lastmodifiedby: ",$propsdoc//cp:lastModifiedBy//text())
                            let $lastmoddate := if(fn:empty($propsdoc//dcterms:modified//text())) then ()
                                                else fn:concat("lastmodified: ",$propsdoc//dcterms:modified//text())
                  return (<a href="#" onclick="openPPTX('{$doc}')">{$doc}</a>,<br/>,
                                <ul class="ML-hit-metadata">
                                     <li>{$lastmodby}</li>&nbsp;&nbsp;
                                     <li>{$lastmoddate}</li>
                                </ul>,<br/>)
return $finaldocs

    
return $return
 
