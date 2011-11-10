xquery version "1.0-ml";
(:
Copyright 2009-2011 MarkLogic Corporation

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
declare namespace ms="http://schemas.openxmlformats.org/spreadsheetml/2006/main";
declare namespace w="http://schemas.openxmlformats.org/wordprocessingml/2006/main";

declare variable $pptadd:bsv as xs:string external;

let $searchparam := $pptadd:bsv

let $w-query := cts:word-query($searchparam)
let $slides := cts:search(/(p:sld|ms:worksheet/ms:sheetData/ms:row|w:document/w:body/w:p), $w-query)
let $uris := for $s in $slides
             return xdmp:node-uri($s)

let $type := for $s at $res in $slides
             let $t := 
               if(fn:not(fn:empty($s/p:cSld))) then
                    let $orig-uri := xdmp:node-uri($s)
                    let $uri := fn:replace($orig-uri,"_pptx_parts/ppt/slides",".pptx")
                       
                    let $tmp-uri := fn:replace($orig-uri,"_pptx_parts/ppt/slides","_PNG")
                    let $tmp-uri2 := fn:replace($tmp-uri,"slide","Slide")
                    let $new-uri := fn:replace($tmp-uri2,".xml",".PNG")
                       
                    let $disp-slides := 
                        for $pic at $d in $new-uri
                        let $src := fn:concat("download-support.xqy?uid=",$pic)
                        let $prop := xdmp:document-properties($pic)
                        let $pptx := $prop/prop:properties/ppt:pptx/text()
                        let $slide := $prop/prop:properties/ppt:slide/text()
                        let $index := $prop/prop:properties/ppt:index/text()

                        let $imageuri := $pic 
                        let $imganchor := fn:concat("#num",$res)
                        let $imgnum := fn:concat("num",$res)
                        return
                        (
                         <div>
                         <li>
                         <table>
                          <tr>
                           <td>
                            <a name={$imgnum} href={$imganchor} onclick="insertSlide('{$pptx}','{$index}','{$res}')">
                             <img src="{$src}" class="resize"></img>
                            </a>
                           </td>
                          </tr>
                         </table>
                              <ul>
                               <li>
                                  <form id={fn:concat("buttons",$res)}>
                                    <input type="radio" name="{$orig-uri}" value="insertslide" id="searchtype" checked="checked"/>Insert Slide
                                    <input type="radio" name="{$orig-uri}" value="embeddocument" id="searchtype" disabled="disabled"/>Embed Document
                                    <input type="radio" name="{$orig-uri}" value="opendocument" id="searchtype"/>Open Document
                                  </form>
                               </li>
                              </ul>
                         </li>
                         <br/>
                         </div>,<br/> 
                        )
                    return <div><ul class="thumb">{$disp-slides}</ul></div>

               else if(fn:not(fn:empty($s/w:r))) then (: its a run of text :)
                    let $snippet := if(string-length(data($s)) > 120) then concat(substring(data($s), 1, 250), "…") else data($s)
                    let $anchor := fn:concat("#num",$res)
                    let $name := fn:concat("num",$res)
                    let $uri := xdmp:node-uri($s)
                    let $text := $s//text()
                    return (<div>
                              <ul>
                               <li title="{data($s)}">
	                          <a name={$name} class="test" href="{$anchor}" onclick="actionDocument('{$res}','{fn:data($s)}')">
                                     {cts:highlight(<p>{$snippet}</p>,$w-query, <strong class="ML-highlight">{$cts:text}</strong>)}
                                  </a>
                               </li>
                               <li>
                                  <form id={fn:concat("buttons",$res)}>
                                    <input type="radio" name="{$uri}" value="inserttext" id="searchtype" checked="checked"/>Insert Text
                                    <input type="radio" name="{$uri}" value="embeddocument" id="searchtype"/>Embed Document
                                    <input type="radio" name="{$uri}" value="opendocument" id="searchtype"/>Open Document
                                  </form>
                               </li>
                              </ul>
                            </div>,
                            <br/>
                           )

               else 
                    let $cells := $s/ms:c
                    let $anchor := fn:concat("#num",$res)
                    let $name := fn:concat("num",$res)
                    let $uri := xdmp:node-uri($s)
                    let $headers := for $hdr in fn:doc($uri)//ms:row[1]/ms:c
                                    return <td class="ML-thdr">{fn:data($hdr)}</td>
                    
                    let $final := for $c in $cells
                                  return <td class="ML-td">{fn:data($c)}</td>
 
                    return <div>
                              <ul>
                                <li>  
                                   {(:<a name={$name} class="test" href="{$anchor}" onclick="openDocument('{$res}')"> :) }
                                   <a name={$name} class="test" href="{$anchor}" onclick="actionDocument('{$res}','{$name}')"> 
                                      <table class="ML-table" id={fn:concat("table",$res)}>
                                                  <tr>{$headers}</tr>
                                                  <tr>{$final}</tr>
                                      </table>
                                   </a>                                 </li>
                                <li>
                                   <form id={fn:concat("buttons",$res)}>
                                      <input type="radio" name="{$uri}" value="inserttable" id="searchtype" checked="checked"/> Insert Table
                                      <input type="radio" name="{$uri}" value="embeddocument" id="searchtype"/>Embed Document
                                      <input type="radio" name="{$uri}" value="opendocument" id="searchtype"/>Open Document
                                   </form>
                                </li>
                              </ul>
                           </div>
                       
               return $t

return $type

