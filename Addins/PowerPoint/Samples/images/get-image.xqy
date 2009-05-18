xquery version "1.0-ml";
let $filename := xdmp:get-request-field("uid")
let $x := fn:doc($filename)
return $x
