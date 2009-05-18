xquery version "1.0-ml";
declare namespace xladd="http://marklogic.com/openxl/exceladdin";
declare variable $xladd:bsv as xs:string external;

(: define variable $searchparam as xs:string external :)
let $searchparam := $xladd:bsv



  let $pics := cts:uri-match(fn:concat("/",$searchparam,"*.jpg"))
  for $pic in $pics
  let $src := fn:concat("get-image.xqy?uid=",$pic)
  let $imageuri := fn:concat("http://localhost:8000/ppt/images/get-image.xqy?uid=",$pic)

  return 
      (<a href="#" onclick="insertImage('{$imageuri}')">
          <img src="{$src}"></img>
       </a>,<br/>,<br/>)
 
