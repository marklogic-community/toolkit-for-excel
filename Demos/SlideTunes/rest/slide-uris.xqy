xquery version "1.0-ml";

import module namespace rest = "http://marklogic.com/appservices/rest"
        at "/MarkLogic/appservices/utils/rest.xqy";

import module namespace json = "http://marklogic.com/json" 
          at "/MarkLogic/appservices/utils/json.xqy";


import module namespace requests =   "http://marklogic.com/appservices/requests" at "requests.xqy";

import module namespace ppt=  "http://marklogic.com/openxml/powerpoint" at "/MarkLogic/openxml/presentation-ml-support.xqy";

let $request := $requests:options/rest:request[@endpoint = "/slide-uris.xqy"][1]

let $map  := rest:process-request($request) 

let $deck := map:get($map, "deck")
let $format := if(fn:empty(map:get($map, "format"))) then
                    "xml"
               else 
                     map:get($map, "format")

let $pptx := fn:tokenize($deck,"/")[last()]
let $ppt-dir := fn:replace($pptx,".pptx","_pptx_parts/ppt/slides/")

let $src-dir := fn:substring-after(fn:substring-before($deck,$pptx),"/office/presentations")
let $directory := fn:concat($src-dir,$ppt-dir)

let $all-uris :=  cts:uris("","document",cts:directory-query($directory,"1"))
let $map := map:map()
let $package:= <slides>{
                   for $u at $i in $all-uris
                   return (<slide><image>{fn:concat($src-dir,$pptx,"/slides/slide",$i)}</image>
                           <single>{fn:data(xdmp:document-properties($u)/prop:properties/ppt:single)}</single></slide>)
                   
                   }</slides>


return if($format eq "json") then
          json:serialize($package)
       else
          $package
