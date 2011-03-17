xquery version "1.0-ml";
(:
Copyright 2008-2011 MarkLogic Corporation

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
declare namespace pkg="http://schemas.microsoft.com/office/2006/xmlPackage";
declare namespace ins="http://marklogic.com/openxml/insert";
declare namespace dc="http://purl.org/dc/elements/1.1/";
declare namespace ppt=  "http://marklogic.com/openxml/excel";

let $uri := xdmp:get-request-field("uri")
let $rId := xdmp:get-request-field("rid")

let $doc-props := xdmp:document-properties($uri)
let $custom-metadata := $doc-props/prop:properties/ppt:shapetags/ppt:shape/ppt:tag[@ppt:rid = $rId]/ppt:custompart

let $metadata := for $meta in $custom-metadata
                 return <meta>{fn:doc($meta)}</meta>

return xdmp:quote(<insertable>
                     {$metadata}
                  </insertable>)

