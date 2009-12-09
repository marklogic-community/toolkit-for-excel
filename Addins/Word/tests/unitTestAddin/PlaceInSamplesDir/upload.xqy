xquery version "1.0-ml"; 
let $filename := xdmp:get-request-field("uid")

let $pkg := xdmp:get-request-body()
let $final :=
    xdmp:document-insert($filename,$pkg) 
return $final

