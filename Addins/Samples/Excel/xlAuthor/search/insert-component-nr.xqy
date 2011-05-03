xquery version "1.0-ml";
(:
Copyright 2008-2011 MarkLogic Corporation

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

declare namespace html = "http://www.w3.org/1999/xhtml";
declare namespace ms = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
declare namespace dc="http://purl.org/dc/elements/1.1/";
import module namespace excel = "http://marklogic.com/openxml/excel" at "/MarkLogic/openxml/spreadsheet-ml-support.xqy";

let $uri := xdmp:get-request-field("uri")
let $start-row-index := xs:integer(xdmp:get-request-field("row")) 
let $start-col-index := xs:integer(xdmp:get-request-field("col"))
let $hit := fn:doc($uri)
let $nrs := $hit/dc:metadata[dc:relation eq "namedrange"]
let $results := for $nr in $nrs
                let $deets:= fn:tokenize($nr/dc:type/node(),"!")
                let $sheetname := $deets[1]
                let $named := $deets[2]
                let $range := (:fn:tokenize("$A$2:$F$7",":"):)fn:tokenize($nr/dc:description[1],":")
                let $s-range := $range[1]
                let $min := fn:tokenize($s-range,"\$")  
                let $min-col := excel:col-letter-to-idx($min[2])
                let $min-row := xs:integer($min[3])
                let $e-range := $range[2]
                let $max := fn:tokenize($e-range,"\$")
                let $max-col := excel:col-letter-to-idx($max[2])
                let $max-row := xs:integer($max[3])
                let $sheet:= fn:doc(fn:concat(fn:substring-before($uri,"customXml"),"xl/worksheets/",fn:lower-case($sheetname),".xml"))
                let $data := $sheet/ms:worksheet/ms:sheetData
                let $delta := $min-row - 1 
                
                let $cells:=   
                   for $cell at $row-idx in ($min-row to $max-row)
                   let $new-row-idx := $start-row-index + $row-idx - 1
                   return 
                                      for $col at $col-idx in $min-col to $max-col                                       
                                      let $new-col-idx := $start-col-index + $col-idx - 1
                                      let $a1 := excel:r1c1-to-a1($row-idx+$delta,$col)
                                      let $cell := $data/ms:row/ms:c[@r=$a1]
                                      return <cell>
                                               <coordinate>{excel:r1c1-to-a1($new-row-idx,$new-col-idx)}</coordinate>
                                               <value>{
                                          if($cell[@t eq "inlineStr"]) then
                                             fn:data($cell/ms:is/ms:t)
                                          else
                                             fn:data($cell/ms:v)
                                              }</value>
                                               <formula>{
                                                fn:data($cell/ms:f)
                                               }</formula>
                                             </cell>
                       
                         return $cells
                
return <insertable>
         {$hit}
         <cells>{$results}</cells>
       </insertable>

