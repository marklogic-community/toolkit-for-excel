xquery version "1.0-ml";

let $uri := xdmp:get-request-field("geturi")
(: let $log := fn:concat("++++++++++++++++++++++++++++++++",$uri) :)
let $x := xdmp:quote(xdmp:http-get($uri)[2]/node())
return $x
