xquery version "1.0-ml";
import module namespace config = "http://marklogic.com/config"  at "./config/config.xqy";
declare namespace search="http://marklogic.com/openxml/search";
xdmp:set-response-content-type('text/html'),
"<!DOCTYPE html PUBLIC '-//W3C//DTD XHTML 1.0 Transitional//EN' 'http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd'>",
<html xmlns="http://www.w3.org/1999/xhtml" id="main">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/> <!--harset=iso-8859-1" />-->
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
   
   <div id="properties-panel"> 
       <h3><span>Properties</span></h3>
       <div id="noproperties">
        No Content Controls are currently selected.
       </div>
       <div id="properties">
       <form action="#">
         <label id="ctrltitle"> </label><br/>
         <label id="ctrltag"> </label><br/>
         <label id="ctrlparent"> </label><br/><br/>
         <input type="checkbox" id="lockctrl" onclick="lockControl()"/><label for="lockctrl">Lock Control</label>
         <input type="checkbox" id="lockcntnt" onclick="lockControlContents()"/><label for="lockcntnt">Lock Content</label>
       </form>
       <br clear="all" />
       </div>
   </div>
   
  </div><!--end inspector details-->
<!-- end following to be generated from config -->
   
  </div><!--end controls-->

  <div id="snippets">
    {config:snippets()}
  </div>
   
</div><!-- end current-doc-->

<div id="metadata">
   <!--going to remove tabs in the future, used here for styling-->
   <!-- info <button onclick="partsTest()">CONTROL TEST</button> -->
   <div id="treeWindow">
      <h2><a href="#">My Document</a></h2>
      <ul id="treelist">
   <!--    <li>test</li>
        <ul><li>test2</li></ul> -->
      </ul>
   </div><!--end treeWindow-->

   <div id="metadataPanel">
    <h3>Metadata</h3>
    <div id="metadataForm">
     <!--<div>
        <p><label>Author</label></p>
        <input id="form1" type="text"/> 
        <p>&nbsp; </p>
     </div>
     <div>
        <p><label>Description</label></p>
        <textarea id="form2"/>
        <p>&nbsp; </p>
     </div>-->
    <!--<p><label>Author</label></p>
    <form id="form1" name="form1" method="post" action="">
      <select name="select" id="select">
        <option>James T. Kirk</option>
      </select>
    </form>
    <p>
      <label>Notes</label>
    </p>
    <form id="form2" name="form2" method="post" action="">
      <textarea name="textarea" cols="40" rows="5" wrap="virtual" id="textarea"></textarea>
    </form> -->
    </div>
   </div><!--end metadataPanel-->

</div>

<div id="search">
   <!--going to remove tabs in the future, used here for styling-->
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
         
               {
                let $res := 
                 if(fn:not(fn:empty($searchparam) or $searchparam eq "" )) then

                    xdmp:invoke("./search/searchresults.xqy",(xs:QName("search:bsv"), $searchparam )) 
                
                 else ()
                 return $res
               }<br/>
               <div id="searchresults">
               </div>
               

	</div>
        { (:if($searchparam eq "" or fn:empty($searchparam)) then $intro else () :)}

 </div>
return $div
}
    {(:xdmp:invoke("searchtest.xqy"):) }
</div>

<div id="compare">
   <!--going to remove tabs in the future, used here for styling-->
  <div id="tabs"></div>
    compare
</div>
</body>
</html>
