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

declare namespace search="http://marklogic.com/openxml/search";
declare namespace w    ="http://schemas.openxmlformats.org/wordprocessingml/2006/main";
declare namespace q    ="http://marklogic.com/beta/searchbox";
declare namespace xlink="http://www.w3.org/1999/xlink";
declare namespace ps = "http://developer.marklogic.com/2006-09-paginated-search";

declare namespace cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
declare namespace dc="http://purl.org/dc/elements/1.1/";
declare namespace dcterms="http://purl.org/dc/terms/";
declare namespace pkg="http://schemas.microsoft.com/office/2006/xmlPackage";


declare variable $LOOKAHEAD-PAGES as xs:integer := 1;
declare variable $SIZE := 10;

declare function ps:page-results(
 $query as cts:query, 
 $start as xs:integer)
 as element(ps:results-page)
{
  let $page-stop := $start + $SIZE -1
  let $stop := 1 + $page-stop + ($SIZE * $LOOKAHEAD-PAGES)
  let $results := cts:search(//w:document//w:sdt, $query)[ $start to $stop ]
  return
    (: if we stepped off of the end, recurse to the previous page :)
    if (empty($results) and ($start - $SIZE) gt 1)
    then ps:page-results($query, $start + $SIZE)
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
          
          let $alias :=  $r/w:sdtPr/w:alias/@w:val
          let $node-uri := xdmp:node-uri($r)

          (:need to check if this came from docx or xml, and adjust accordingly :)
          let $doc-uri := if(fn:contains($node-uri, "_docx_parts")) then
                              fn:concat(fn:substring-before ($node-uri,"_docx_parts"),".docx")
                          else
                              $node-uri
          let $docprops := if(fn:contains($node-uri, "_docx_parts")) then
                                fn:doc(fn:concat(fn:substring-before($node-uri,"word/document.xml"),"docProps/core.xml"))
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
           attribute path { xdmp:path($r) },
           attribute title { $title },
           attribute ctrltype { $alias },
           attribute modby { $last-mod-by },
           attribute moddate { $last-mod-date },
           $r
          }
      
      }

};

declare function ps:get-element-attribute-value-query($params) as cts:query?
{
   let $queries := for $p in $params 
                   return cts:element-attribute-value-query(
                     xs:QName("w:alias"),
                     xs:QName("w:val"),
                      $p) 
   return cts:or-query($queries)   
  
};

declare function ps:get-and-query($raw as xs:string, $params as xs:string*)
 as cts:query?
{ let $param := if(fn:empty($params)) then () else ps:get-element-attribute-value-query($params)
  let $words := tokenize($raw, '\s+')[. ne '']
  where $words
  return cts:and-query(($param,$words))
};

declare function ps:get-or-query($raw as xs:string, $params as xs:string*)
 as cts:query?
{
  let $words := tokenize($raw, '\s+')[. ne '']
  where $words
  return cts:or-query($words)
};


let $q := xdmp:get-request-field("qry")
let $params := xdmp:get-request-field("params")
let $new-start := if(fn:empty(xdmp:get-request-field("start"))) then 
                    1 
                  else 
                      xs:integer(xdmp:get-request-field("start"))
let $intro := 
       <div id="ML-Intro">
	<h1>Search and Reuse for {$q}</h1>
	<p>Use the above search box to find content in Word 2007 documents stored on MarkLogic Server. Keywords narrow the results. Each search result represents a paragraph (or list item) that matches your criteria.</p>
	<p>To insert that entire paragraph into the active document at the current cursrrror location, double-click the result snippet.</p>
       </div>
(:return xdmp:quote($intro) :)
return	xdmp:quote(
          if($q) then
            let $and-query := ps:get-and-query($q,$params)
            let $or-query := ps:get-or-query($q,$params)
	    let $tokens := tokenize($q, "\s+")
            let $hits := ps:page-results($and-query, $new-start)
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
                         if(fn:not($hits) or fn:empty($hits//w:sdt)) then
	                       (<div id="searchresultsinner"><p>Your search for "{$q}" did not match anything.</p>{$intro}</div>)
		         else
                               (   
                                <div id="searchpagination">
                                <span>Results:</span>
                                {
 
                                    if($remainder gt 10) then 
                                    let $page := $new-start
                                    let $new-page := $new-start + 10
                                    return if($page gt 10) then
                                             (<a href="#" class="leftpagination" OnClick="SearchAction({$page - 10})">&lt;</a>,
                                                 $span,
                                              <a href="#" class="rightpagination" OnClick="SearchAction({$new-page})">&gt;</a>)
                                           else 
                                              ("&nbsp;&nbsp;",$span,<a href="#" class="rightpagination" OnClick="SearchAction({$new-page})">&gt;</a>)
                                                                   
                                     else if($new-start gt 10) then
                                             (<a href="#" class="leftpagination" OnClick="SearchAction({$new-start - 10})">&lt;</a>,
                                              $span
                                             )
                                     else  $span
                                } 
                                                        
                                </div>,
                                
                                
                                for $hit in $hits/ps:result 
				let $uri := fn:data($hit/@uri)
                                let $path := fn:data($hit/@path)
(: title,ctrltype , modby, moddate :)
                                let $ctrl := $hit/w:sdt 
				let $snippet := if(string-length(fn:data($ctrl)) > 120) then 
                                                   fn:concat(substring(fn:data($ctrl), 1, 120), "...") 
                                                else fn:data($ctrl)
				return 
                                 <div class="searchreturnresult">
                                 <h4>{fn:data($hit/@title)}</h4>
                                 <p class="byline">{fn:concat("Modified: ",fn:data($hit/@moddate))}&nbsp; 
                                          <span>{fn:data($hit/@modby)}</span>
                                 </p>
                                 <p class="controltitle">
                                     <span class="textIcon">{fn:data($hit/@ctrltype)}</span>
                                     <!--<span class="pagenum">p. 63</span>-->
                                 </p> 
				     {cts:highlight(<p class="searchreturnsnippet" title="{fn:data($ctrl)}">{$snippet}</p>, $or-query, <strong class="ML-highlight">{$cts:text}</strong>)}
					 <div id="searchresultactions">
					    <!--<a href="./utils/content.xqy?uri={xdmp:url-encode($uri)}" target="_blank">INSERT</a>&nbsp;-->
					    <a href="#" class="insertbtn" OnClick="InsertAction('{xdmp:url-encode($uri)}', '{$path}')">INSERT</a>&nbsp;
					    <a href="./utils/openpkg.xqy?uri={xdmp:url-encode($uri)}" class="openbtn">OPEN</a>
                                         </div>
                                 </div>
                               ) } </div>
           return ($res)
                                                                   
         else ()
         )

 
