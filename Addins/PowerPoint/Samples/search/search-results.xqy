xquery version "1.0-ml";
declare namespace xladd="http://marklogic.com/openxl/exceladdin";
declare namespace html = "http://www.w3.org/1999/xhtml";
declare namespace a="http://schemas.openxmlformats.org/drawingml/2006/main";
declare namespace r="http://schemas.openxmlformats.org/officeDocument/2006/relationships";
declare namespace p="http://schemas.openxmlformats.org/presentationml/2006/main";
declare namespace dc = "http://purl.org/dc/elements/1.1/";
declare namespace cp = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
declare namespace dcterms="http://purl.org/dc/terms/";

declare variable $xladd:bsv as xs:string external;
declare variable $xladd:searchtype as xs:string external;

let $searchparam := $xladd:bsv
let $searchtype :=  $xladd:searchtype

let $return := 
if($searchtype eq "slide") then
     let $slides := cts:search(//p:sld, cts:word-query($searchparam))
     let $slideuris := for $s in $slides 
                       let $orig-uri := xdmp:node-uri($s)
                       (:let $tmp-uri := fn:replace($orig-uri,"_parts/ppt/slides","_parts_GIF"):)
                       let $tmp-uri := fn:replace($orig-uri,"_pptx_parts/ppt/slides","_GIF")
                       let $tmp-uri2 := fn:replace($tmp-uri,"slide","Slide")
                       let $new-uri := fn:replace($tmp-uri2,".xml",".GIF")
                       return $new-uri

   (:  let $properties := xdmp:document-properties($slideuris)
     let $pptx := $properties//pptx/text()
     let $slide := $properties//slide/text()
     let $index := $properties//index/text() :)
     
     (: let $slides := cts:uri-match(fn:concat("/*",$searchparam,"*.GIF")) :)
     let $disp-slides := 
         for $pic at $d in $slideuris
         let $src := fn:concat("download-support.xqy?uid=",$pic)
         (:let $imageuri := fn:concat("http://localhost:8023/ppt/search/get-image.xqy?uid=",$pic) :)
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
else if($searchtype eq "image") then
     let $pics := cts:uri-match(fn:concat("/",$searchparam,"*.jpg"))
     for $pic at $d in $pics
       let $src := fn:concat("download-support.xqy?uid=",$pic)
       let $imganchor := fn:concat("#num",$d)
       let $imgnum := fn:concat("num",$d) 

       (:construct the url string in js, using config from Addin for url:)
       let $imageuri := $pic (: fn:concat("http://localhost:8023/ppt/search/get-image.xqy?uid=",$pic)  :)
       return 
         (<a name={$imgnum} href={$imganchor} onclick="insertImage('{$imageuri}')">
          <img src="{$src}"></img>
          </a>,<br/>,<br/>)
else
let $slides := cts:search(//p:sld, cts:word-query($searchparam))
let $docuris := for $s in $slides 
                       let $orig-uri := xdmp:node-uri($s)
                       return fn:replace(fn:replace($orig-uri,"/ppt/slides/slide\d+\.xml",""),"_pptx_parts",".pptx")
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
 
