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

declare namespace xladd="http://marklogic.com/openxl/exceladdin";
declare variable $xladd:bsv as xs:string external;

(: define variable $searchparam as xs:string external :)
let $searchval := if(fn:empty($xladd:bsv) or $xladd:bsv eq "") then () else $xladd:bsv
let $header:=((: xdmp:set-response-content-type('text/html'), :)
              <div id="header">
                 <form id="basicsearch" action="default.xqy" method="post">
                   <div>
                      <input type="text" size="40" name="xladd:bsv" autocomplete="off" value="{$searchval}" id="bsearchval"  method="post"/>&nbsp;
                     <!-- TEST : { $no:color}--><input type="submit" value="Search"/> 
                       
                   </div> 
                  </form>    
             </div>)

return $header
