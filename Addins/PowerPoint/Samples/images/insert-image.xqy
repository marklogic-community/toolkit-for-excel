xquery version "1.0-ml";
(:
Copyright 2009-2011 MarkLogic Corporation

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
	<script type="text/javascript" src="../MarkLogicPowerPointAddin.js">//</script>

	<script type="text/javascript" src="images.js">//</script>
</head>
{
let $rgb :=  "rgb(200,216,237)"
let $searchparam := if(fn:empty(xdmp:get-request-field("xladd:bsv"))) then "" else (xdmp:get-request-field("xladd:bsv"))
let $body :=
      <body bgcolor={$rgb}>
	<div id="ML-Add-in">
<br/>
               {
                    xdmp:invoke("image-search.xqy",  (xs:QName("xladd:bsv"),$searchparam ))
               }
               <br/><br/>
            
               {
                let $res := 
                 if(fn:not(fn:empty($searchparam) or $searchparam eq "" )) then

                     xdmp:invoke("image-results.xqy",(xs:QName("xladd:bsv"), $searchparam ))
                
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
