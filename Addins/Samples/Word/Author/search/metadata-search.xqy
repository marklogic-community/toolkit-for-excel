xquery version "1.0-ml";
(:
Copyright 2008-2010 Mark Logic Corporation

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
:)

declare namespace search="http://marklogic.com/openxml/search";
declare namespace w    ="http://schemas.openxmlformats.org/wordprocessingml/2006/main";
declare namespace q    ="http://marklogic.com/beta/searchbox";
declare namespace xlink="http://www.w3.org/1999/xlink";
declare namespace ps = "http://developer.marklogic.com/2006-09-paginated-search";

declare namespace cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
declare namespace dc="http://purl.org/dc/elements/1.1/";
declare namespace dcterms="http://purl.org/dc/terms/";
declare namespace pkg="http://schemas.microsoft.com/office/2006/xmlPackage";

let $q := xdmp:get-request-field("qry")
(:let $docname := "SITE1.docx":)
let $results := cts:search(//dc:metadata,$q)
(:let $node := cts:uri-match("*SITE1.docx"):)

let $uris :=<options id="docnames">
             {(:assumes coming from extracted folder, check for saved as .xml :)
              for $r in $results
              let $doc-folder := fn:substring-before(xdmp:node-uri($r), "/customXml/item") 
              let $doc-uri := fn:replace($doc-folder,"_docx_parts", ".docx")
              let $doc-name := fn:tokenize($doc-uri,"/")[last()]
              return <option value="{fn:concat($doc-folder,"/")}">{$doc-name}</option>
             }
             </options>

(:let $uris :=  for $r in $results
              return fn:replace(fn:substring-before(xdmp:node-uri($r), "/customXml/item"),"_docx_parts", ".docx"):)
return xdmp:quote($uris)

