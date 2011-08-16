xquery version "1.0-ml";

import module namespace rest = "http://marklogic.com/appservices/rest"
        at "/MarkLogic/appservices/utils/rest.xqy";

import module namespace json = "http://marklogic.com/json" 
          at "/MarkLogic/appservices/utils/json.xqy";

import module namespace requests =   "http://marklogic.com/appservices/requests" at "requests.xqy";

declare variable $offset := 9;

let $request := $requests:options/rest:request[@endpoint = "/playlists.xqy"][1]

let $map  := rest:process-request($request) 

let $directory := map:get($map, "directory")

let $format := if(fn:empty(map:get($map, "format"))) then
                    "xml"
               else 
                     map:get($map, "format")

let $start := if(fn:empty(map:get($map, "start"))) then 
                  0 
              else
                  map:get($map,"start")

let $end := $start + $offset


(: way to filter .pptx in query? :)
let $uris :=  cts:uris("","document",cts:directory-query($directory,"infinity"))
(: or use mimetype instead:)

let $package := <playlists>{
                   for $f in $uris[$start to $end]
                   return <playlist>{$f}</playlist>
                 }</playlists>

return if($format eq "json") then
          json:serialize($package)
       else
          $package 
