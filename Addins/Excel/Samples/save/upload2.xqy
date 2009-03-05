xquery version "1.0-ml"; 
let $filename := xdmp:get-request-field("uid")
let $pkg := xdmp:get-request-body()
 return
  xdmp:document-insert(fn:concat("/",$filename),$pkg) 

