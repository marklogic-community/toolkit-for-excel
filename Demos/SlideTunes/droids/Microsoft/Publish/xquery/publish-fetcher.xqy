xquery version "1.0-ml";
declare namespace  ppt = "http://marklogic.com/openxml/powerpoint";
declare variable $doc as xs:string external;

let $doc := fn:doc($doc)
return $doc/playlist/slides/slide/single/text()

