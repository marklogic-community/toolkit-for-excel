xquery version "1.0-ml";

import module namespace rest = "http://marklogic.com/appservices/rest"
        at "/MarkLogic/appservices/utils/rest.xqy";


import module namespace requests =   "http://marklogic.com/appservices/requests" at "requests.xqy";

let $request := $requests:options/rest:request[@endpoint = "/directory-delete.xqy"][1]

let $map  := rest:process-request($request) 

let $directory := map:get($map, "directory")

return (xdmp:directory-delete($directory),"SUCCESS")
