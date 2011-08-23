xquery version "1.0-ml";
declare namespace html = "http://www.w3.org/1999/xhtml";
declare namespace foo="http://marklogic.com/foo";


let $uri := xdmp:get-request-field("uri")
let $plname := xdmp:get-request-field("plname")
let $puturi := fn:concat($uri, $plname)

let $gallery := xdmp:get-request-field("gallery")
let $xml-gallery:= xdmp:unquote($gallery)

let $log := fn:concat("URI:",$uri," PLNAME", $plname," Gallery",$gallery)

(:let $publish :=
       <ppt:publish>{
       for $x in $xml-gallery/gallery/source
       return xdmp:document-properties(fn:string($x))/prop:properties/ppt:single
       }</ppt:publish>

return xdmp:document-insert($uri,$publish)

:)
(:
let $put := xdmp:http-put($puturi, (), $xml-gallery)
return $put
:)
return xdmp:document-insert($plname, $xml-gallery)


