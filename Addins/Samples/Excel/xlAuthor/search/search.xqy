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

declare namespace search="http://marklogic.com/openxml/search";
declare namespace w    ="http://schemas.openxmlformats.org/wordprocessingml/2006/main";
declare namespace q    ="http://marklogic.com/beta/searchbox";
declare namespace a="http://schemas.openxmlformats.org/drawingml/2006/main";
declare namespace r="http://schemas.openxmlformats.org/officeDocument/2006/relationships";
declare namespace rel ="http://schemas.openxmlformats.org/package/2006/relationships";
declare namespace p="http://schemas.openxmlformats.org/presentationml/2006/main";
declare namespace xlink="http://www.w3.org/1999/xlink";
declare namespace ps = "http://developer.marklogic.com/2006-09-paginated-search";
declare namespace cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
declare namespace dc="http://purl.org/dc/elements/1.1/";
declare namespace dcterms="http://purl.org/dc/terms/";
declare namespace pkg="http://schemas.microsoft.com/office/2006/xmlPackage";
declare namespace ms = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

import module namespace excel = "http://marklogic.com/openxml/excel" at "/MarkLogic/openxml/spreadsheet-ml-support.xqy";

import module namespace ppt=  "http://marklogic.com/openxml/powerpoint" at "/MarkLogic/openxml/presentation-ml-support.xqy";

declare variable $LOOKAHEAD-PAGES as xs:integer := 1;
declare variable $SIZE := 10;

declare function ps:page-results(
 $query as cts:query, 
 $search-type as xs:string,
 $start as xs:integer)
 as element(ps:results-page)
{
  let $page-stop := $start + $SIZE -1
  let $stop := 1 + $page-stop + ($SIZE * $LOOKAHEAD-PAGES)
  let $results := if($search-type eq "workbook") then
                      cts:search(/ms:worksheet/ms:sheetData/ms:row, $query)[ $start to $stop ]
                  else (: if($search-type eq "macro") then :)
                      cts:search(/dc:metadata[dc:relation eq "macro"], $query)[ $start to $stop ]
  return
    (: if we stepped off of the end, recurse to the previous page :)
    if (empty($results) and ($start - $SIZE) gt 1)
    then ps:page-results($query, $search-type,$start + $SIZE)
    else
      let $count := count($results)
      let $remainder :=
        if (exists($results)) then cts:remainder($results[1]) else 0
      let $estimated := $remainder gt $count
      return element ps:results-page {
        attribute estimated { $estimated },
        attribute remainder { max(( $remainder, $count )) },
        attribute start { $start },
          for $r in $results[1 to $page-stop]
          
         
          let $node-uri := xdmp:node-uri($r)

          (:need to check if this came from docx or xml, and adjust accordingly :)
          let $doc-uri := if(fn:contains($node-uri, "_xlsx_parts")) then
                              fn:concat(fn:substring-before ($node-uri,"_xlsx_parts"),".xlsx")
                          else if (fn:contains($node-uri, "_xlsm_parts")) then
                              fn:concat(fn:substring-before ($node-uri,"_xlsm_parts"),".xlsm")
                          else
                              $node-uri

          let $docprops := if(fn:contains($node-uri, "_xlsx_parts")) then
                                fn:doc(fn:concat(fn:substring-before($node-uri,"xl/worksheets"),"docProps/core.xml"))
                           else if(fn:contains($node-uri, "_xlsm_parts")) then
                                 fn:doc(fn:concat(fn:substring-before($node-uri,"customXml"),"docProps/core.xml"))
                           else
                                fn:doc($node-uri)/pkg:package/pkg:part[@pkg:name="/docProps/core.xml"]/pkg:xmlData

                                 (: need fn:doc of pkg part containing core.xml :)
          let $last-mod-by := $docprops/cp:coreProperties/cp:lastModifiedBy/text()
          let $date :=  $docprops/cp:coreProperties/dcterms:created/text()
          let $title :=if(fn:empty($docprops/cp:coreProperties/dc:title/text())) then
                         $doc-uri 
                       else
                          $docprops/cp:coreProperties/dc:title
                          
          let $last-mod-date := fn:concat(fn:month-from-dateTime (xs:dateTime($date))
                 ,"/",
                  fn:day-from-dateTime (xs:dateTime($date))
                 ,"/",
                  fn:year-from-dateTime (xs:dateTime($date))
                 ,"  ",
                  fn:hours-from-dateTime (xs:dateTime($date))
                  ,":",
                  fn:hours-from-dateTime (xs:dateTime($date))
                 )

          return  element ps:result {
           attribute uri { $node-uri },
           attribute docuri { $doc-uri },
           attribute path { xdmp:path($r) },
           attribute title { $title },

           attribute modby { $last-mod-by },
           attribute moddate { $last-mod-date },
           ( $r/preceding-sibling::*[1],$r,$r/following-sibling::*[1])
          }
      
      }

};

declare function ps:get-cts-property-query($params as xs:string*, $search-type as xs:string*) as cts:query?
{
   let $queries := for $p in $params 
                   let $elem-val-query := if($search-type eq "component") then
                                           cts:element-query(xs:QName("ppt:shapetags"),  cts:element-value-query(xs:QName("ppt:tagname"),$p) )
                                          else if ($search-type eq "slide") then
                                           (cts:element-query(xs:QName("ppt:slidetags"),  cts:element-value-query(xs:QName("ppt:tagname"),$p) ),
                                             cts:element-query(xs:QName("ppt:presentationtags"),  cts:element-value-query(xs:QName("ppt:tagname"),$p))
                                           )
                                          else
                                           cts:element-query(xs:QName("ppt:presentationtags"),  cts:element-value-query(xs:QName("ppt:tagname"),$p) )

                   return cts:properties-query($elem-val-query)
   return cts:or-query($queries)   
  
};

declare function ps:get-and-query($raw as xs:string, $params as xs:string*, $search-type as xs:string*)
 as cts:query?
{ (: let $param := if(fn:empty($params)) then () else ps:get-cts-property-query($params, $search-type) :)
  let $words := tokenize($raw, '\s+')[. ne '']
  where $words
  return cts:and-query($words)
};

declare function ps:get-or-query($raw as xs:string, $params as xs:string*)
 as cts:query?
{
  let $words := tokenize($raw, '\s+')[. ne '']
  where $words
  return cts:or-query($words)
};



let $q := xdmp:get-request-field("qry")
let $search-type := xdmp:get-request-field("stype")

let $params := xdmp:get-request-field("params")
let $new-start := if(fn:empty(xdmp:get-request-field("start"))) then 
                    1 
                  else 
                      xs:integer(xdmp:get-request-field("start"))
let $intro := 
       <div id="ML-Intro">
	<h2>Search and Reuse</h2>
	<p>Use the above search box to find content in Excel 2007 documents stored on MarkLogic Server. Keywords narrow the results. Each search result represents a component or document that matches your criteria.</p>
        <br/>
	<p>To insert the results into the active workbook on the current worksheet selection, click the insert button.  To open the source document for the search result, click the document title.  Mouseover the snippet or image to see more detail about the search result.</p>
       </div>
return	xdmp:quote(
          if($q) then
            let $and-query := ps:get-and-query($q,$params, $search-type)
            let $or-query := ps:get-or-query($q,$params)
            let $hits := ps:page-results($and-query, $search-type, $new-start)
            let $remainder := fn:data($hits/@remainder)
            let $new-end := if($remainder gt $new-start+9) then $new-start + 9 else $new-start + $remainder - 1
            let $span := <span class="resultscounter">{$new-start} to {$new-end} of 
                          {
                            if(fn:data($hits/@remainder) gt $new-end) then 
                                fn:data($hits/@remainder)
                            else $new-end
                          }
                          </span> 
  	    let $res := <div>
                        {
                         if(fn:not($hits) or (fn:empty($hits//ms:c) and fn:empty($hits//dc:metadata))) then
	                       (<div id="searchresultsinner">
                                     <p>Your search for "{$q}" did not match anything.</p>
                                     {$intro}
                                </div>)
		         else
                               (   
                                <div id="searchpagination">
                                <strong>Results:</strong>
                                { (: Results area, pagination for $span :)
 
                                    if($remainder gt 10) then 
                                    let $page := $new-start
                                    let $new-page := $new-start + 10
                                    return if($page gt 10) then
                                             (<a href="javascript:searchAction({$page - 10});" class="leftpagination">
                                                    <img src="images/arrow-small-left.png"  border="0"/>
                                              </a>,
                                              $span,
                                              <a href="javascript:searchAction({$new-page});" class="rightpagination">
                                                 <img src="images/arrow-small-right.png"  border="0"/>
                                              </a>)
                                           else 
                                              ("&nbsp;&nbsp;",
                                               $span,
                                               <a href="javascript:searchAction({$new-page});" class="rightpagination"> 
                                                 <img src="images/arrow-small-right.png"  border="0"/>
                                               </a>)
                                                                   
                                     else if($new-start gt 10) then
                                             (<a href="javascript:searchAction({$new-start - 10});" class="leftpagination">
                                                 <img src="images/arrow-small-left.png"  border="0"/>
                                              </a>,
                                              $span
                                             )
                                     else  ("&nbsp;","&nbsp;",$span)
                                } 
                                </div>,
                                for $hit at $idx in $hits/ps:result 
				let $uri := fn:data($hit/@uri)
                                let $doc-uri := fn:data($hit/@docuri)
                                let $sheet-num := fn:substring-before(fn:substring-after($uri,"parts/xl/worksheets/"),".xml")

                                let $slideuri :=xdmp:document-properties($uri)/prop:properties/ppt:slideimg/text()
                                let $src := fn:concat("search/download-support.xqy?uid=",$slideuri)

  (:determine tags here :)
                                let $tag-rid := if(fn:local-name($hit/p:sp) eq "sp") then 
                                                   fn:data($hit/p:sp//p:tags/@r:id)
                                                else if (fn:local-name($hit/p:pic) eq "pic") then 
                                                   fn:data($hit/p:pic//p:tags/@r:id)
                                                else
                                                   fn:data($hit/p:sld/p:cSld/p:custDataLst/p:tags/@r:id)

                                let $ctrl := if(fn:local-name($hit/p:sp) eq "sp") then 
                                                 $hit/p:sp 
                                             else if (fn:local-name($hit/p:pic) eq "pic") then 
                                                 $hit/p:pic
                                             else
                                                 $hit/ms:worksheet

                                let $pic :=  if(fn:local-name($hit/p:pic) eq "pic") then
                                               let $pic-rid := $hit/p:pic//a:blip/@r:embed
                                               let $s-rels := fn:data($hit/@sliderels)
                                               let $img-uri := fn:replace(fn:data(fn:doc($s-rels)//rel:Relationship[@Id = $pic-rid]/@Target),"\.\.","")
                                               let $ppt-dir := fn:substring-before($s-rels,"/slides/")
                                               let $img := fn:concat($ppt-dir, $img-uri)
                                               return $img                                          
                                             else 
                                                 ""

                                let $rows := $hit/ms:row
                                let $cells := $hit/ms:row/ms:c
                                let $anchor := fn:concat("#num",$idx)
                                let $headers := for $hdr in fn:doc($uri)//ms:row[1]/ms:c
                                                return <td class="ML-thdr">{fn:data($hdr)}</td>
                    
                                let $final := for $row in $rows return <tr>{for $c in $row/ms:c
                                              return <td class="ML-td">{fn:data($c)}</td>}</tr>

                                let $snippet :=  <table class="ML-table" id={fn:concat("table",$idx)}>
                                                   <tr>{$headers}</tr>
                                                   <tr>{$final}</tr>
                                                 </table>


                                let $icon-type := if (fn:local-name($hit/p:pic) eq "pic") then
                                                      "imageIcon"
                                                  else if(fn:local-name($hit/p:sp) eq "sp") then
                                                      "textIcon"
                                                  else "slideIcon"

				return 
                                 <div class="searchreturnresult">
                                  <h4>{ if($search-type eq "workbook") then 
                                          <a href="./utils/openpkg.xqy?uri={xdmp:url-encode($doc-uri)}" onmouseup="blurSelected(this)" class="blacklink">
                                            {fn:data($hit/@title)}  <!-- presentation name  -->
                                          </a>
                                        else
                                             fn:data($hit/@title)
                                       }
                                  </h4> 
                                 <p class="byline">{fn:concat("Modified: ",fn:data($hit/@moddate))}&nbsp; 
                                          <span>{fn:data($hit/@modby)}</span> <!-- presentation metadata  -->
                                 </p>

<!-- if workbook/sheet, list of workbooks only, make title link to open, if opened -->
<!-- if macro, list of macros, with buttons for insert, run, undo -->
<!-- if component , then need to send json as well as metadata -->
				 {
                                  let $page := if($search-type eq "workbook") then
                                                 let $highlight := cts:highlight(
                                                                                $snippet , 
                                                                                $or-query, 
                                                                                <strong class="ML-highlight">{$cts:text}</strong>) 
                                                 let $sheetdeet :=  <p class="byline">{$sheet-num}</p>
                                                 return ($sheetdeet, $highlight)
                                               else if ($search-type eq "macro") then
                                                 let $macro-name := fn:data($hit/dc:metadata/dc:identifier[1])
                                                 let $macro-text := fn:string($hit/dc:metadata/dc:description[1])
                                                 let $macro-desc := <p class="searchreturnsnippet" title="{$macro-text}"> {$hit/dc:metadata/dc:description[2]}</p>
                                                 let $action :=  <div id="searchresultactions">
                                                                   <a href="javascript:insertMacroAction('{xdmp:url-encode($uri)}','{$idx}');" onmouseup="blurSelected(this)" class="smallbtn searchinsertbtn"> 
                                                                     <span>&nbsp;&nbsp;Add</span>
                                                                   </a>
                                                                   <a href="javascript:runMacro('{$macro-name}');" onmouseup="blurSelected(this)" class="smallbtn searchrunbtn">
                                                                     <span id="{fn:concat("runbutton",$idx)}">&nbsp;&nbsp;&nbsp;Run</span>
                                                                   </a>&nbsp;&nbsp;
                                                                   <a href="javascript:removeMacro('{$macro-name}');" onmouseup="blurSelected(this)" class="smallbtn"> 
                                                                        <span id="{fn:concat("removebutton",$idx)}">
                                                                               Remove 
                                                                        </span>
                                                                   </a>
                                                                   
                                                                 </div>
                                                 return ($macro-desc, $action)
                                               else (: its a component :)
                                                 let $highlight := cts:highlight(<p class="searchreturnsnippet" title="{fn:data($ctrl)}">
                                                                                    <span class="{$icon-type}">&nbsp;</span>
                                                                                       {$snippet} </p>, 
                                                                                       $or-query, 
                                                                                     <strong class="ML-highlight">{$cts:text}</strong>)  
                                                  let $img := <img src="{$src}" class="resize"></img>
                                                      let $action :=  <div id="searchresultactions">
                                            {
                                              if ($search-type eq "component") then
					        <a href="javascript:insertComponentAction('{xdmp:url-encode($uri)}', '{$tag-rid}','{$pic}','{$idx}');" onmouseup="blurSelected(this)" class="smallbtn searchinsertbtn"><span>Insert</span></a>
                                              else 
                                                   <a href="javascript:insertSlideAction('{xdmp:url-encode($uri)}', '{$tag-rid}','{xdmp:url-encode($doc-uri)}','{$idx}');" onmouseup="blurSelected(this)" class="smallbtn searchinsertbtn"><span>Insert</span></a>
					    }
                                              &nbsp;
<!-- <a href="./utils/openpkg.xqy?uri={xdmp:url-encode($uri)}" class="smallbtn">Open</a> -->
                                            <!-- can probably reuse delete used on first pane -->
					    <span id="{fn:concat("undobutton",$idx)}"><a href="javascript:undoInsert();" onmouseup="blurSelected(this)" class="smallbtn">Undo</a></span>
                                                                       </div>
                                                       return ($highlight, $img,<br/>, $action)
                                        return $page
                                      }
                                 </div>
                               ) } </div>
           return ($res)
                                                                   
         else ()
         )

 
