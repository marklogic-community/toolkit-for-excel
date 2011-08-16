xquery version "1.0-ml";

import module namespace rest = "http://marklogic.com/appservices/rest"
        at "/MarkLogic/appservices/utils/rest.xqy";

import module namespace json = "http://marklogic.com/json" 
          at "/MarkLogic/appservices/utils/json.xqy";

import module namespace requests =   "http://marklogic.com/appservices/requests" at "requests.xqy";

let $request := $requests:options/rest:request[@endpoint = "/playlist-fetch.xqy"][1]

let $map  := rest:process-request($request) 

let $deck := map:get($map, "deck")

let $fullpath := fn:substring-after($deck,"/playlists")
let $filename := fn:tokenize($deck,"/")[last()]
let $json := fn:ends-with($fullpath,".json")

return if($json) then
          let $doc := fn:doc(fn:replace($fullpath,".json",".xml"))
          return json:serialize($doc/node())
         
       else
          fn:doc($fullpath)
