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
import module namespace config = "http://marklogic.com/config"  at "../config/config.xqy";
declare namespace search = "http://marklogic.com/openxml/search";
declare variable $search:bsv as xs:string external;

(: define variable $searchparam as xs:string external :)
let $searchval := if(fn:empty($search:bsv) or $search:bsv eq "") then "Search..." else $search:bsv
let $header:=((: xdmp:set-response-content-type('text/html'), :)
              <div id="searchhead" xmlns="http://www.w3.org/1999/xhtml"><!-- adds line -->
                 <!--<form id="basicsearch" action="default.xqy" method="get">-->
                   <div id="searchform">
                      <input type="text" id="searchbox" style="width:100px" value="{$searchval}" onkeypress="checkForEnter();"/>
                      
		      <!--<input id="ML-Submit" type="submit" value="search"/>-->
                     
                      <a href="#"  id="sbtn" class="searchbtn" onmouseup="blurSelected(this);" onclick="SearchAction();" >Search</a>
                      <a href="#" id="fbtn" class="filterbtn"> Filter </a>
                      <a href="#" id="ddbtn" class="dropdownbtn"> &nbsp; </a>
                      <!-- in mockup, but not available yet
                     <a href="#" class="filterBtn">&nbsp;</a>
                     <a href="#" class="dropdownBtn">&nbsp;</a>
                        -->
		      
                      <!--<input type="text" name="search:bsv" autocomplete="off" value={$searchval} id="bsearchval"  method="post"/>
                      <input type="submit" value="search"/> -->
                       
                   </div><br clear="all" /><br clear="all" />
<!-- this needs to be a configurable query, and we probably won't do counts in v1, what does facet/link mean, since its a filter? -->
                      {config:search-filters()}
                    <!--<div id="searchfilter">
                        <div class="filterrow"><input type="checkbox" id="Section" /><a href="#"> Section</a></div>
                        <div class="filterrow"><input type="checkbox" id="Policy" /><a href="#"> Policy</a></div>
                        <div class="filterrow"><input type="checkbox" id="Process" /><a href="#"> Process </a></div>
                   </div>-->
 
                 <!-- </form>    -->
             </div>)

return $header
