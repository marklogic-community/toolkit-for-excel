xquery version "1.0-ml";
import module namespace config = "http://marklogic.com/config"  at "./config/config.xqy";
xdmp:set-response-content-type('text/html'),
"<!DOCTYPE html PUBLIC '-//W3C//DTD XHTML 1.0 Transitional//EN' 'http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd'>",
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link rel="stylesheet" href="css/sample1.css" />
<script type="text/javascript" src="jquery-1.3.2.min.js"></script>
<script type="text/javascript" src="sample1.js"></script>
<script type="text/javascript" src="MarkLogicWordAddin.js">//</script>
<script type="text/javascript" src="MarkLogicContentControlSupport.js">//</script>

<!-- generate a lookuptable in js as well , so I can add appropriate metadata for control-->
<script type="text/javascript">{config:generate-js-for-controls()}</script>

<!-- some dynamic script here for insert of controls -->
<title>Oslo Information Panel</title>
</head>
<body> 
<div id="topnav">
	<ul>
    	<li><a href="#" class ="fronticon" id="icon-word">&nbsp;</a></li>
        <li><a href="#" id="icon-metadata">&nbsp;</a></li>
        <li><a href="#" id="icon-search">&nbsp;</a></li>
        <li><a href="#" id="icon-merge">&nbsp;</a></li>
     </ul>
    <br clear="all" />
</div>

<div id="current-doc">
  <div id="tabs">
   <ul>
     <li class="fronttab" id="controltab"><a href="#" id="controls-show">Controls</a></li>
     <li id="snippettab"><a href="#" id="snippets-show">Snippets</a></li>
   </ul>
  </div>
  <div id="controls">
     <div id="buttonbar">
	<ul>
    	  <li><a href="#" class="selectedctrl" id="icon-textctrl">&nbsp;</a></li>
          <li><a href="#" id="icon-imgctrl">&nbsp;</a></li>
          <li><a href="#" id="icon-calctrl">&nbsp;</a></li>
          <li><a href="#" id="icon-dropctrl">&nbsp;</a></li>
          <li><a href="#" id="icon-comboctrl">&nbsp;</a></li>
        </ul>
      </div>

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

   <!-- <button onclick="partsTest()">GENERAL TEST BUTTON</button> -->


   </div><!--end text controls-->
   <div id="imgcontrols">
       <h3><span>Insert picture  control</span></h3>
       <ul class="buttongroup">
         {config:picctrl-inline()}
       </ul>
       <br clear="all" />
   </div>
   <div id="calcontrols">
       <h3><span>Insert calendar control</span></h3>
       <ul class="buttongroup">
         {config:calctrl-inline()}
       </ul>
       <br clear="all" />
   </div>
   <div id="dropcontrols">
       <h3><span>Insert dropdown control</span></h3>
       <ul class="buttongroup">
         {config:dropctrl-inline()}
       </ul>
       <br clear="all" />
   </div>
   <div id="combocontrols">
       <h3><span>Insert combo control</span></h3>
       <ul class="buttongroup">
         {config:comboctrl-inline()}
       </ul>
       <br clear="all" />
   </div>
   
   <div id="properties"> 
       <h3><span>Properties</span></h3>
       <form action="#">
         <label id="ctrltitle"> </label><br/>
         <label id="ctrltag"> </label><br/>
         <input type="checkbox" id="lockctrl" onclick="lockControl()"/><label for="lockctrl">Lock Control</label>
         <input type="checkbox" id="lockcntnt" onclick="lockControlContents()"/><label for="lockcntnt">Lock Content</label>
       </form>
       <br clear="all" />
   </div>
   
  </div><!--end inspector details-->
<!-- end following to be generated from config -->
   
  </div><!--end controls-->

  <div id="snippets">
    {config:snippets()}
  </div>
   
</div><!-- end current-doc-->

<div id="metadata">
  <div id="tabs"></div>
    info
</div>

<div id="search">
  <div id="tabs"></div>
    search 
</div>

<div id="compare">
  <div id="tabs"></div>
    compare
</div>
</body>
</html>
