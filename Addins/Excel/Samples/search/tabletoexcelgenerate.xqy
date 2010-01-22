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

declare namespace excel = "http://marklogic.com/openxml/excel";
import module "http://marklogic.com/openxml/excel" at "/MarkLogic/openxml/spreadsheet-ml-support.xqy";
declare namespace html = "http://www.w3.org/1999/xhtml";
declare namespace ms = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
declare namespace r="http://schemas.openxmlformats.org/officeDocument/2006/relationships";
declare namespace xladd="http://marklogic.com/openxl/exceladdin";

let $wrapperelem := xdmp:get-request-field("elemname") 
let $docuri := xdmp:get-request-field("docuri")
let $doc := "fn:doc($docuri)"

let $path := fn:concat($doc,"//",$wrapperelem)

let $testname := $wrapperelem 
let $tabstyle := xs:boolean("true")
let $colcustwidths := "25" 

let $wbname := if($testname eq "") then "Default"
                  else if(fn:empty($testname)) then "Default"
                  else $testname

let $xlsxname := fn:concat($wbname,".xlsx")
let $original := xdmp:unpath($path)
let $valid1 := excel:validate-child($original)

let $final := if($valid1 eq xs:boolean("false")) then
(
xdmp:set-response-content-type('text/html;charset=utf-8'),
'<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">',
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="content-type" content="text/html; charset=utf-8" />
	<title>Create Worksheet from XML Table</title>
	<link rel="stylesheet" type="text/css" href="../css/office-blue.css"/>
	<script type="text/javascript" src="../MarkLogicExcelAddin.js">//</script>
	<script type="text/javascript" src="search.js">//</script>
</head>
{
let $rgb :=  "rgb(200,216,237)"
let $body :=
      <body bgcolor={$rgb}> 
      <br/>
      <img src="ackbar2.jpg"/>
<div id="ML-Add-in">
<!-- <div id="ML-Message"> -->
     <!-- <h1>ACK! It's A Trap!!!</h1> -->
     <p><strong>ACK! It's A Trap!!!</strong></p>
<br/><br/>
       
     <p> That doesn't appear to be a table,<br/> 
      Please try to open another document.</p>
      <br/><br/>
      <a href="default.xqy?xladd:bsv={$wrapperelem}">Go Back</a>
<!--</div> -->
</div>
	<div id="ML-Navigation">
	   <a href="../default.xqy">« Samples</a>
        </div>
      </body>
return $body
}
</html>
) 
else
(
    let $package := excel:create-xlsx-from-xml-table($original,$colcustwidths,fn:true(),$tabstyle)

    let $filename := $xlsxname 
    let $disposition := concat("attachment; filename=""",$filename,"""")
    let $x := xdmp:add-response-header("Content-Disposition", $disposition)
    let $x := xdmp:set-response-content-type("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet") 
    return
      $package
)

return $final
