xquery version "1.0-ml";
declare namespace xladd="http://marklogic.com/openxl/exceladdin";
declare variable $xladd:bsv as xs:string external;
declare variable $xladd:searchtype as xs:string external;

let $searchparam := $xladd:bsv
let $searchtype :=  $xladd:searchtype

let $return := 
if($searchtype eq "slide") then
     let $slides := cts:uri-match(fn:concat("/*",$searchparam,"*.GIF"))
     let $disp-slides := 
         for $pic in $slides
         let $src := fn:concat("download-support.xqy?uid=",$pic)
         (:let $imageuri := fn:concat("http://localhost:8023/ppt/search/get-image.xqy?uid=",$pic) :)
         let $imageuri := $pic 
         return
          (<li><a href="#" onclick="copyPasteSlideToActive('{$imageuri}')">
             <img src="{$src}" class="resize"></img>
           </a></li> )
     return <div>{$searchtype}<br/><ul class="thumb">{$disp-slides}</ul></div>
else
     let $pics := cts:uri-match(fn:concat("/",$searchparam,"*.jpg"))
     for $pic in $pics
       let $src := fn:concat("download-support.xqy?uid=",$pic)

       (:construct the url string in js, using config from Addin for url:)
       let $imageuri := $pic (: fn:concat("http://localhost:8023/ppt/search/get-image.xqy?uid=",$pic)  :)
       return 
         (<a href="#" onclick="insertImage('{$imageuri}')">
          <img src="{$src}"></img>
          </a>,<br/>,<br/>)
return $return
 
