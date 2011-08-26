xquery version "1.0-ml";

let $uri := xdmp:get-request-field("uri")
let $log := fn:concat("++++++++++++++++++++++++++++++++",$uri) 
return xdmp:http-delete($uri) 
