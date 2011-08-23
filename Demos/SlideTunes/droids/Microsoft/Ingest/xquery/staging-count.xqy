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
(: let $filename := xdmp:get-request-field("uri")
return xdmp:quote(fn:doc($filename)) :)

(: cts:uris("",(),cts:directory-query("/staging/","1")) :)

let $uris := cts:uris("",(),cts:directory-query("/staging/","1"))
let $to-process :=
       for $u in $uris
       return if(xdmp:document-properties($u)/prop:properties/status) then ()
              else $u
let $active := 
       for $t in $to-process
       return xdmp:document-set-property($t,<status>active</status>)

return $to-process

