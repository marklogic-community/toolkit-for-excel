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

declare namespace w="http://schemas.openxmlformats.org/wordprocessingml/2006/main";
declare namespace xladd="http://marklogic.com/openxl/exceladdin";
declare namespace q    ="http://marklogic.com/beta/searchbox";

xdmp:set-response-content-type('text/html;charset=utf-8'),
(:'<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">',:)
'<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">',
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="content-type" content="text/html; charset=utf-8" />
	<title>Create Worksheet from XML Table</title>
	<link rel="stylesheet" type="text/css" href="../css/office-blue.css"/>
	<script type="text/javascript" src="../js/MarkLogicExcelAddin.js">//</script>
	<script type="text/javascript" src="search.js">//</script>
</head>
{
let $rgb :=  "rgb(200,216,237)"
let $intro :=   <div id="ML-Intro">
			<h1>Search and Reuse</h1>
			<p>Use the above search box to find content in Excel 2007 documents and XML documents that contain tabular data stored on MarkLogic Server.
				Keywords narrow the results. Each search result represents a document that matches your criteria.</p>
			<p>Click the document title to open into Excel. If the document is an Excel document, it will open right up.  If it's XML, an attempt will be made to convert that XML to an Excel Worksheet and open into Excel for you. </p>
		</div>

let $searchparam := if(fn:empty(xdmp:get-request-field("xladd:bsv"))) then "" else (xdmp:get-request-field("xladd:bsv"))
let $body :=
      <body bgcolor={$rgb}>
	<div id="ML-Add-in">
<br/>
               {
                    xdmp:invoke("xlsearch.xqy",  (xs:QName("xladd:bsv"),$searchparam ))
               }
               <br/><br/>
            
               {
                let $res := 
                 if(fn:not(fn:empty($searchparam) or $searchparam eq "" )) then

                     xdmp:invoke("xlresults.xqy",(xs:QName("xladd:bsv"), $searchparam ))
                
                 else ()
                 return $res
               }<br/>
               

	</div>
        { if($searchparam eq "" or fn:empty($searchparam)) then $intro else () }

	<div id="ML-Navigation">
	   <a href="../default.xqy">« Samples</a>
        </div>
 </body>
return $body
}
</html>
