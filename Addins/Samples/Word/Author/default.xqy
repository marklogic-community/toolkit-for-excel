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
import module namespace config = "http://marklogic.com/config"  at "./config/config.xqy";
declare namespace search="http://marklogic.com/openxml/search";
xdmp:set-response-content-type('text/html'),
"<!DOCTYPE html PUBLIC '-//W3C//DTD XHTML 1.0 Transitional//EN' 'http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd'>",
<html xmlns="http://www.w3.org/1999/xhtml" id="main">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/> <!--harset=iso-8859-1" />-->
<link rel="stylesheet" href="css/authoring.css" />
<script type="text/javascript" src="js/jquery-1.3.2.min.js"></script>
<script type="text/javascript" src="js/authoring.js"></script>
<script type="text/javascript" src="js/MarkLogicWordAddin.js">//</script>
<script type="text/javascript" src="js/MarkLogicContentControlSupport.js">//</script>

<!-- generate a lookuptable in js as well , so I can add appropriate metadata for control-->
<script type="text/javascript">{config:generate-js-for-controls()}</script>

<!-- some dynamic script here for insert of controls -->
<title>Oslo Information Panel</title>
</head>
<body> 
  <div id="topnav">
	<ul>
    	<li><a href="#" class ="fronticon" id="icon-word" title="enrich">&nbsp;</a></li>
        <li><a href="#" id="icon-metadata" title="metadata">&nbsp;</a></li>
        <li><a href="#" id="icon-search" title="search">&nbsp;</a></li>
        <li><a href="#" id="icon-merge" title="compare">&nbsp;</a></li>
     </ul>
    <br clear="all" />
  </div><!--end topnav-->

  <div id="current-doc">
    <div id="tabs">
     <ul>
       <li class="fronttab" id="controltab" title="control palette">
            <a href="#" id="controls-show">Controls</a>
       </li>
       <li id="snippettab">
            <a href="#" id="snippets-show" title="insert boilerplate">Boilerplate</a>
       </li>
     </ul>
    </div><!-- end tabs-->
    <div id="controls">
       <div id="buttonbar">
	  <ul>
    	    <li><a href="#" class="selectedctrl" id="icon-textctrl" title="rich text">&nbsp;</a></li>
            <li><a href="#" id="icon-imgctrl" title="image">&nbsp;</a></li>
            <li><a href="#" id="icon-calctrl" title="calendar">&nbsp;</a></li>
            <li><a href="#" id="icon-dropctrl" title="dropdown">&nbsp;</a></li>
            <li><a href="#" id="icon-comboctrl" title="combobox">&nbsp;</a></li>
          </ul>
        </div><!-- end buttonbar-->

  <!-- begin following to be generated from config -->
        <div class="inspectorDetails">
          <div id="textcontrols">
            <h3><span>Insert section control</span></h3>
            <ul class="buttongroup">
              {config:textctrl-sections()}
            </ul>
            <br clear="all" />
            <h3><span>Insert inline control</span></h3>
            <ul class="buttongroup">
              {config:textctrl-inline()}
            </ul>
            <br clear="all" />
          </div><!--end text controls-->
          <div id="imgcontrols">
            <h3><span>Insert picture  control</span></h3>
            <ul class="buttongroup">
              {config:picctrl-inline()}
            </ul>
            <br clear="all" />
          </div> <!--end imgcontrols-->
          <div id="calcontrols">
            <h3><span>Insert calendar control</span></h3>
            <ul class="buttongroup">
              {config:calctrl-inline()}
            </ul>
            <br clear="all" />
          </div><!-- end calcontrols-->
          <div id="dropcontrols">
            <h3><span>Insert dropdown control</span></h3>
            <ul class="buttongroup">
              {config:dropctrl-inline()}
            </ul>
            <br clear="all" />
          </div><!--end dropcontrols-->
          <div id="combocontrols">
            <h3><span>Insert combo control</span></h3>
            <ul class="buttongroup">
              {config:comboctrl-inline()}
            </ul>
            <br clear="all" />
          </div><!--end combocontrols-->
   
          <div id="properties-panel"> 
            <h3><span>Properties</span></h3>
            <div id="noproperties">
               No Content Controls are currently selected.
            </div>
            <div id="properties">
              <form action="#">
                <strong><label id="ctrltitle"></label></strong>
                <span style="color:#458CBB"> | </span>
                <label id="ctrltag"> </label>
                <span style="color:#458CBB"> | </span>
                <label id="ctrlparent"> </label>
                <br/><br/>
                <input type="checkbox" id="lockctrl" onclick="lockControl()"/><label for="lockctrl">Lock Control</label><br/>
                <input type="checkbox" id="lockcntnt" onclick="lockControlContents()"/><label for="lockcntnt">Lock Content</label>
              </form>
              <br clear="all" />
            </div><!--end properties-->
          </div><!--end properties-panel-->
        </div><!--end inspector details-->
<!-- end following to be generated from config -->
    </div><!--end controls-->

    <div id="snippets">
      <div id="boilerplate" class="inspectorDetails">
        <h3>Insert boilerplate content</h3>
        <ul class="buttongroup">
          {config:snippets()}
        </ul>
      </div>
    </div><!--end snippets-->
  </div><!-- end current-doc-->

  <div id="metadata">
    <div id="treeWindow">
      <h2><a href="#">My Document</a></h2>
      <ul id="treelist">
      </ul>
    </div><!--end treeWindow-->

    <div id="metadataPanel">
      <h3>Metadata</h3>
      <div id="metadataForm">
      </div>
    </div><!--end metadataPanel-->
  </div><!--end metadata-->

  <div id="search">
  { 
    let $searchparam := if(fn:empty(xdmp:get-request-field("search:bsv"))) then "" else (xdmp:get-request-field("search:bsv"))
    let $start := xdmp:get-request-field("search:start")
    let $div :=
      <div id="searchpanel">
	<div>
         <br/>
               {
                    xdmp:invoke("./search/searchform.xqy",  (xs:QName("search:bsv"),$searchparam ))
               }
         
               <div id="searchresults">
               </div>

	</div>
      </div>
    return $div
  }
  </div><!-- end search-->

  <div id="compare" style="display:block">
    <div id="mergemenuarea">Location
      <ul class="menu">
        <li>
          <a  onClick="displayLayer('ulmenu')" class="menu" id="select0">-choose-</a>
            {config:compare-filters()}
        </li>
      </ul>
    </div><!-- end of mergemenuarea --> 
  </div><!-- end of compare -->

  <div id="mergesection" style="display:block">
    <div id="mergeresults">
    <!-- CHANGE DOCNAMES TO SOMETHING ELSE!! -->
      <div id="docnames">
       </div>
    </div>
  </div><!--end mergesection-->
</body>
</html>
