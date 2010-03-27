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

declare namespace search = "http://marklogic.com/openxml/search";
declare variable $search:bsv as xs:string external;

(: define variable $searchparam as xs:string external :)
let $searchval := if(fn:empty($search:bsv) or $search:bsv eq "") then () else $search:bsv
let $header:=((: xdmp:set-response-content-type('text/html'), :)
              <div id="searchheader">
                 <!--<form id="basicsearch" action="default.xqy" method="get">-->
                   <div>
                      <input id="ML-Search" name="search:bsv" autocomplete="off" type="text" value="{$searchval}" method="get"/>
                      
		      <!--<input id="ML-Submit" type="submit" value="search"/>-->
                     
                      <a href="#" OnClick="SearchAction();" > Search </a> &nbsp;|&nbsp;
                      <a href="#"> Filter </a>
		      
                      <!--<input type="text" name="search:bsv" autocomplete="off" value={$searchval} id="bsearchval"  method="post"/>
                      <input type="submit" value="search"/> -->
                       
                   </div> 
                 <!-- </form>    -->
             </div>)

return $header
