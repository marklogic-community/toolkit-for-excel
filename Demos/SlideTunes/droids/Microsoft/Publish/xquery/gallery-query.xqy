xquery version "1.0-ml";

let $uris := cts:uris("",(),cts:directory-query("/gallery/","1"))

(:let $to-process :=
       for $u in $uris
       return if(xdmp:document-properties($u)/prop:properties/status) then ()
              else $u
let $active := 
       for $t in $to-process
       return xdmp:document-set-property($t,<status>active</status>):)

return $uris

