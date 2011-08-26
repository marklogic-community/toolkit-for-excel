xquery version "1.0-ml";
declare namespace html = "http://www.w3.org/1999/xhtml";
declare namespace foo="http://marklogic.com/foo";


let $plname := xdmp:get-request-field("plname")
let $simplename := fn:tokenize($plname,"/")[last()]
let $publishname := fn:concat("/publish/",$simplename)
let $filename := fn:concat("/out/",fn:replace($simplename,".xml",".pptx"))

let $playlist := fn:doc($plname)
let $insert := xdmp:document-insert($publishname, $playlist)
return $filename


