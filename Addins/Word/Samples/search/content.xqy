xquery version "1.0-ml";
(:
Copyright 2008-2009 Mark Logic Corporation

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

import module namespace xp = "http://marklogic.com/xinclude/xpointer" at "/MarkLogic/xinclude/xpointer.xqy";
declare namespace w="http://schemas.openxmlformats.org/wordprocessingml/2006/main";

let $href := xdmp:url-decode(xdmp:get-request-field('uri'))
let $tokens := tokenize($href, "#")
let $uri := $tokens[1]
let $ptr := $tokens[2]
return (
	if(ends-with($uri, ".docx")) then (
			xdmp:set-response-content-type('application/vnd.openxmlformats-officedocument.wordprocessingml.document'),
			doc($uri)
	)	else (
		xdmp:set-response-content-type('application/xml'),
		if($ptr) then
			xp:dereference(doc($uri), $ptr)
		else 
			doc($uri)
	)
)
