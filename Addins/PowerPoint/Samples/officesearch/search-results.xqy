xquery version "1.0-ml";
declare namespace xladd="http://marklogic.com/openxl/exceladdin";
declare namespace html = "http://www.w3.org/1999/xhtml";
declare namespace a="http://schemas.openxmlformats.org/drawingml/2006/main";
declare namespace r="http://schemas.openxmlformats.org/officeDocument/2006/relationships";
declare namespace p="http://schemas.openxmlformats.org/presentationml/2006/main";
declare namespace dc = "http://purl.org/dc/elements/1.1/";
declare namespace cp = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
declare namespace dcterms="http://purl.org/dc/terms/";
declare namespace ms="http://schemas.openxmlformats.org/spreadsheetml/2006/main";
declare namespace w="http://schemas.openxmlformats.org/wordprocessingml/2006/main";

declare variable $xladd:bsv as xs:string external;

let $searchparam := $xladd:bsv

let $w-query := cts:word-query($searchparam)
let $slides := cts:search(/(p:sld|ms:worksheet/ms:sheetData/ms:row|w:document/w:body/w:p), $w-query)
let $uris := for $s in $slides
             return xdmp:node-uri($s)

let $type := for $s at $res in $slides
             let $t := if(fn:not(fn:empty($s/p:cSld))) then
                        let $orig-uri := xdmp:node-uri($s)
                       
                        let $tmp-uri := fn:replace($orig-uri,"_pptx_parts/ppt/slides","_GIF")
                        let $tmp-uri2 := fn:replace($tmp-uri,"slide","Slide")
                        let $new-uri := fn:replace($tmp-uri2,".xml",".GIF")
                       
                       let $disp-slides := 
                        for $pic at $d in $new-uri
                        let $src := fn:concat("download-support.xqy?uid=",$pic)
                        let $prop := xdmp:document-properties($pic)
                        let $pptx := $prop//pptx/text()
                        let $slide := $prop//slide/text()
                        let $index := $prop//index/text()

                        let $imageuri := $pic 
                        let $imganchor := fn:concat("#num",$d)
                        let $imgnum := fn:concat("num",$d)
                        return
                        (
                <div>
<li>
                <table>
                  <tr>
                    <td><a name={$imgnum} href={$imganchor} onclick="copyPasteSlideToActive('{$pptx}','{$index}','{$d}')">
                          <img src="{$src}" class="resize"></img>
                        </a>
                    </td>
                    <td style="vertical-align: top;" >
                          <input type="checkbox" id={fn:concat("retain",$d)} name="format"/>retain format
                    </td>
                 </tr>
                </table>
           </li>
              <br/>
              </div>,<br/> 
                        )
                        return <div>{(:$searchtype:)}<ul class="thumb">{$disp-slides}</ul></div>
    

                       else if(fn:not(fn:empty($s/w:r))) then
                       let $snippet := if(string-length(data($s)) > 120) then concat(substring(data($s), 1, 250), "…") else data($s)
                       let $anchor := fn:concat("#num",$res)
                       let $name := fn:concat("num",$res)
                       let $uri := xdmp:node-uri($s)
                       let $text := $s//text()
                       return (<div>
                               <ul>
                                <li title="{data($s)}">
	                           <a name={$name} class="test" href="{$anchor}" onclick="openWord('{$res}','{fn:data($s)}')">
                                     {cts:highlight(<p>{$snippet}</p>,$w-query, <strong class="ML-highlight">{$cts:text}</strong>)}
                                   </a>
                                </li>
                                <li>
                            <form id={fn:concat("buttons",$res)}>
                            <input type="radio" name="{$uri}" value="inserttext" id="searchtype"/>Insert Text
                            <input type="radio" name="{$uri}" value="opendocument" id="searchtype"/>Open Document
                            </form>
                                </li>
                              </ul>
                              </div>,<br/>)

                       else 
                       let $cells := $s/ms:c
                       let $anchor := fn:concat("#num",$res)
                       let $name := fn:concat("num",$res)
                       let $uri := xdmp:node-uri($s)
                       let $final := for $c in $cells
                                          return <td class="ML-td">{fn:data($c)}</td>
 
                       return <div>
                              <ul>
                                <li>  
                                   <a name={$name} class="test" href="{$anchor}" onclick="openWord('{$res}')">
                                      <table class="ML-table"><tr>{$final}</tr></table>
                                   </a>
                                </li>
                                <li>
                                  <form id={fn:concat("buttons",$res)}>
                                  <input type="radio" name="{$uri}" value="embeddocument" id="searchtype"/>Embed Document
                                  <input type="radio" name="{$uri}" value="opendocument" id="searchtype"/>Open Document
                                  </form>
                                </li>
                              </ul>
                              </div>
                       
             return $t

return $type

