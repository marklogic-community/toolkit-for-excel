xquery version "1.0-ml";

import module namespace rest = "http://marklogic.com/appservices/rest"
        at "/MarkLogic/appservices/utils/rest.xqy";

import module namespace requests =   "http://marklogic.com/appservices/requests" at "requests.xqy";

let $request := $requests:options/rest:request[@endpoint = "/slide-image.xqy"][1]

let $map  := rest:process-request($request) 

let $deck := fn:substring-after(map:get($map, "deck"),"/office/presentations")

let $tokens := fn:tokenize($deck,"/")
let $pptx := $tokens[last()-1]

let $directory := fn:substring-before($deck,$pptx)

let $slide := map:get($map, "slide")
let $size := if(fn:empty(map:get($map, "size"))) then 
                  "small" 
             else
                  map:get($map,"size")

let $suffix := if($size eq "small") then
                  "_BMP_S"
               else if($size eq "medium") then
                  "_BMP_M"
               else if($size eq "large") then
                  "_BMP_L"
               else
                  "_BMP_S"
                  
let $image-dir := fn:concat($directory,fn:replace($pptx,".pptx",$suffix),"/")

let $slide-img := fn:concat($image-dir,"Slide",$slide,".BMP")
return 
  fn:doc($slide-img)
