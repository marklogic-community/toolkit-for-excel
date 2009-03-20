xquery version "1.0-ml";

declare namespace excel = "http://marklogic.com/openxml/excel";
import module "http://marklogic.com/openxml/excel" at "/MarkLogic/openxml/excel-ml-support.xqy";
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
let $worksheetname := $original/fn:local-name(child::*[1])
(: let $worksheetrows := //submissions/child::*[1] :)
let $worksheetrows := $original/child::*

(: order of rows?  :)
let $allrows := for $r at $d in $worksheetrows
             let $rowhdrs := $r//child::*
             let $validrownames :=
                    for $i in $rowhdrs 
                    let $rowhdrname := fn:local-name($i)
                    return $rowhdrname
             return $validrownames
let $headerrows :=  fn:distinct-values($allrows)
let $columncount := fn:count($headerrows)

let $headers := excel:create-row($headerrows)

let $rowvalues := for $i at $d in $worksheetrows
                  let $map := map:map()
                  let $return := for $x at $z in $i/child::*
                                 let $put := map:put($map, fn:local-name($x),$x/text())
                                 return $put 
                  return $map

let $rows := for $i in $rowvalues
             return excel:create-row($i,$headerrows) 

let $rowcount := fn:count($rows)
            
let $content-types := excel:create-simple-content-types(1,xs:boolean("true"))
let $workbook := excel:create-simple-workbook(1)
let $rels :=  excel:create-simple-pkg-rels()
let $workbookrels :=  excel:create-simple-workbook-rels(1)

let $tablerange := fn:concat("A1:",excel:r1c1-to-a1($rowcount+1,$columncount))
let $styling := $tabstyle (: if($tabstyle eq "mark1") then xs:boolean("true") else xs:boolean("false") :)
let $tablexml :=  excel:create-simple-table($tablerange, $headerrows, $styling)

let $worksheetrels := excel:create-simple-worksheet-rels()
let $sheet-col-widths := for $i in 1 to $columncount return $colcustwidths (: ("14","58","16","16","18","24") :)
let $colwidths := excel:worksheet-cols($sheet-col-widths) 

let $sheet1 := excel:create-simple-worksheet(($headers,$rows), $colwidths, xs:boolean("true")) 

let $package := excel:generate-simple-xl-pkg($content-types, $workbook, $rels, $workbookrels, $sheet1, $worksheetrels, $tablexml)

    let $filename := $xlsxname 
    let $disposition := concat("attachment; filename=""",$filename,"""")
    let $x := xdmp:add-response-header("Content-Disposition", $disposition)
    let $x := xdmp:set-response-content-type("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet") 
    return
      $package
)

return $final
