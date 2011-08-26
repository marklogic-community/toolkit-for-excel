xquery version "1.0-ml";


let $filename := xdmp:get-request-field("filename")
return fn:exists(fn:doc($filename))


