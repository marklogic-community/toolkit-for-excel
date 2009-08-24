xquery version "1.0-ml";
(:
Copyright 2009 Mark Logic Corporation

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

declare namespace w="http://schemas.openxmlformats.org/wordprocessingml/2006/main";
declare namespace xladd="http://marklogic.com/openxl/exceladdin";

xdmp:set-response-content-type('text/html;charset=utf-8'),
(:'<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">',:)
'<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">',
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="content-type" content="text/html; charset=utf-8" />
	<title>Insert Images</title>
	<link rel="stylesheet" type="text/css" href="../css/office-blue.css"/>
	<link rel="stylesheet" type="text/css" href="slides.css"/>
	<script type="text/javascript" src="../MarkLogicPowerPointAddin.js">//</script>
	<script type="text/javascript" src="jquery-1.3.2.js">//</script>

	<script type="text/javascript" src="search.js">//</script>
</head>
{
let $rgb :=  "rgb(200,216,237)"
let $searchparam := if(fn:empty(xdmp:get-request-field("xladd:bsv"))) then "" else (xdmp:get-request-field("xladd:bsv"))
let $searchtype :=  if(fn:empty(xdmp:get-request-field("xladd:searchtype"))) then "" else (xdmp:get-request-field("xladd:searchtype"))
(:do test for checked here :)
let $searchval := $searchparam (: if(fn:empty($xladd:bsv) or $xladd:bsv eq "") then () else $xladd:bsv :)
let $body :=
      <body bgcolor={$rgb}>
	<div id="ML-Add-in">
<br/>
<form id="basicsearch" action="default.xqy" method="post">
                   <div>
                      <input type="text" size="40" name="xladd:bsv" autocomplete="off" value={$searchval} id="bsearchval"  method="post"/>&nbsp;
                     <!-- TEST : { $no:color}--><input type="submit" value="Go"/> 
                  </div>     
                   <br/> {
                         if($searchtype eq "slide")
                         then
                            <input type="radio" name="xladd:searchtype" checked="checked" value="slide" id="s"/>
                         else
                            <input type="radio" name="xladd:searchtype" value="slide" id="s"/>
                         }Slides
                         {
                         if($searchtype eq "image")
                         then
                            <input type="radio" name="xladd:searchtype" value="image" checked="checked" id="i"/>
                         else
                            <input type="radio" name="xladd:searchtype" value="image" id="i"/>
                         }Images
                         {
                         if($searchtype eq "pres")
                         then
                            <input type="radio" name="xladd:searchtype" value="pres" checked="checked" id="i"/>
                         else
                            <input type="radio" name="xladd:searchtype" value="pres" id="i"/>
                         }Presentations
                  </form>   
               {(:
                    xdmp:invoke("image-search.xqy",  (xs:QName("xladd:bsv"),$searchparam ))
               :)}
               <br/><br/>
            
               {
                let $res := 
                 if(fn:not(fn:empty($searchparam) or $searchparam eq "" )) then

                     xdmp:invoke("search-results.xqy",((xs:QName("xladd:bsv"), $searchparam),(xs:QName("xladd:searchtype"), $searchtype)))
                
                 else ()
                 return $res
               }<br/>
               

	</div>
	<div id="ML-Navigation">
	   <a href="../default.xqy">« Samples</a>
        </div>
 </body>
return $body
}
</html>
