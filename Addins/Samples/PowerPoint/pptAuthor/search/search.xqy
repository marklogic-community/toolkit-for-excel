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
  let $results := if($search-type eq "component") then 
                      let $elem-query := cts:element-query(xs:QName("p:tags"), cts:and-query(()))
                      let $search := cts:search(//p:sld, cts:and-query(($elem-query,$query)))[ $start to $stop ]
                      let $results := for $res in $search//p:sp//p:tags/ancestor::p:sp | $search//p:pic//p:tags/ancestor::p:pic
                                      return $res
                      return $results
(: fn:data($hit/p:sp//p:tags/@r:id) :)
                      

                  else if($search-type eq "slide") then
                      cts:search(//p:sld, $query)[ $start to $stop ]
                  else 
                       let $search := cts:search(//p:sld, $query)
                       let $docuris := for $s in $search 
                                       let $orig-uri := xdmp:node-uri($s)
                                       return xdmp:document-properties($orig-uri)/prop:properties/ppt:pptxdir/text()
                       let $final := for $d in  fn:distinct-values($docuris)[ $start to $stop ]
                                     return fn:doc(fn:concat($d,"ppt/slides/slide1.xml")) (:need to do better than this :)
                       return $final
                                     
                 
  return
    (: if we stepped off of the end, recurse to the previous page :)
    if (empty($results) and ($start - $SIZE) gt 1)
    then ps:page-results($query, $search-type, $start + $SIZE)
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

          let $slide-idx := fn:replace(fn:replace(fn:substring-after($node-uri,"ppt/slides/"),"slide",""),".xml","" )
   
          let $slide-rels := fn:replace($node-uri,"slide\d+\.xml",fn:concat("_rels/slide",$slide-idx,".xml.rels"))

          (:need to check if this came from docx or xml, and adjust accordingly :)
          let $doc-uri := if(fn:contains($node-uri, "_pptx_parts")) then
                              fn:concat(fn:substring-before ($node-uri,"_pptx_parts"),".pptx")
                          else
                              $node-uri

          let $docprops := if(fn:contains($node-uri, "_pptx_parts")) then
                                fn:doc(fn:concat(fn:substring-before($node-uri,"ppt/slides"),"docProps/core.xml"))
                                (:fn:doc(fn:concat(fn:substring-before($node-uri,"ppt/presentation.xml"),"docProps/core.xml")):)                           else
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
           attribute docuri { $doc-uri },
           attribute sliderels { $slide-rels },
           (: attribute ctrltype { $alias }, :)
           (: previously, control type was control label, so use tag label 
              we also need type for icon display                     :)
           (: using this function for both slides and components, so putting logic for components below :)
           attribute modby { $last-mod-by },
           attribute moddate { $last-mod-date },
           $r
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
{ let $param := if(fn:empty($params)) then () else ps:get-cts-property-query($params, $search-type) 
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
let $search-type := xdmp:get-request-field("stype")

let $params := xdmp:get-request-field("params")
let $new-start := if(fn:empty(xdmp:get-request-field("start"))) then 
                    1 
                  else 
                      xs:integer(xdmp:get-request-field("start"))
let $intro := 
       <div id="ML-Intro">
	<h2>Search and Reuse</h2>
	<p>Use the above search box to find content in PowerPoint 2007 documents stored on MarkLogic Server. Keywords narrow the results. Each search result represents a component or document that matches your criteria.</p><br/>
	<p>To insert the results into the active presentation on the current slide selection, click the insert button.  To open the source document for the search result, click the open button.  Mouseover the snippet or image to see more detail about the search result.</p>
       </div>
return	xdmp:quote(
          if($q) then
            let $and-query := ps:get-and-query($q,$params, $search-type)
            let $or-query := ps:get-or-query($q,$params)
	    let $tokens := tokenize($q, "\s+")
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
                         (: if(fn:not($hits) or fn:empty($hits//w:sdt)) then :)
                         if(fn:not($hits) or ( fn:empty($hits//p:sld) and fn:empty($hits//p:pic) and fn:empty($hits//p:sp) )) then
	                       (<div id="searchresultsinner"><p>Your search for "{$q}" did not match anything.</p>{$intro}</div>)
		         else
                               (   
                                <div id="searchpagination">
                                <strong>Results:</strong>
                                { (: Results area, pagination for $span :)
 
                                    if($remainder gt 10) then 
                                    let $page := $new-start
                                    let $new-page := $new-start + 10
                                    return if($page gt 10) then
                                             (<a href="javascript:searchAction({$page - 10});" class="leftpagination"> <img src="images/arrow-small-left.png"  border="0"/></a>,
                                                 $span,
                                              <a href="javascript:searchAction({$new-page});" class="rightpagination"> <img src="images/arrow-small-right.png"  border="0"/></a>)
                                           else 
                                              ("&nbsp;&nbsp;",$span,<a href="javascript:searchAction({$new-page});" class="rightpagination"> <img src="images/arrow-small-right.png"  border="0"/></a>)
                                                                   
                                     else if($new-start gt 10) then
                                             (<a href="javascript:searchAction({$new-start - 10});" class="leftpagination"> <img src="images/arrow-small-left.png"  border="0"/></a>,
                                              $span
                                             )
                                     else  ("&nbsp;","&nbsp;",$span)
                                } 
                                                        
                                </div>,
(: below needs fixing :)
(: change to look at param, not rid:)
                              (: let $new-hits := if($search-type eq "component") then 
                                                 for $hit at $idx in $hits/ps:result
                                                 let $ret := for $p in $params
                                                             let $hit-rid := if(fn:local-name($hit/p:sp) eq "sp") then 
                                                                                fn:data($hit/p:sp//p:tags/@r:id)
                                                                             else if (fn:local-name($hit/p:pic) eq "pic") then 
                                                                                fn:data($hit/p:pic//p:tags/@r:id)
                                                                             else
                                                                                 fn:data($hit/p:sld/p:cSld/p:custDataLst/p:tags/@r:id)
                                                             let $final := if(xdmp:document-properties($hit)//ppt:shapetags/ppt:shape/ppt:tag/@ppt:rid eq $hit-rid) then $p else () 
                                                             return $final
                                                 return $ret
                                                else
                                                   $hits
(: above needs fixing :)
                                 :)
                                
                                
                                for $hit at $idx in $hits/ps:result 
				let $uri := fn:data($hit/@uri)
                                let $pptxuri := fn:data($hit/@docuri)

                                let $slide-idx := fn:replace(fn:replace(fn:substring-after($uri,"ppt/slides/"),"slide",""),".xml","" )


                                let $slideuri :=xdmp:document-properties($uri)/prop:properties/ppt:slideimg/text()

                                let $src := fn:concat("search/download-support.xqy?uid=",$slideuri)

                                let $path := fn:data($hit/@path)
                                (:  do check here on type, could be p:sld or p:sp :)

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
                                                 $hit/p:sld

                                let $pic :=  if(fn:local-name($hit/p:pic) eq "pic") then
                                               let $pic-rid := $hit/p:pic//a:blip/@r:embed
                                               let $s-rels := fn:data($hit/@sliderels)
                                               let $img-uri := fn:replace(fn:data(fn:doc($s-rels)//rel:Relationship[@Id = $pic-rid]/@Target),"\.\.","")
                                               let $ppt-dir := fn:substring-before($s-rels,"/slides/")
                                               let $img := fn:concat($ppt-dir, $img-uri)
                                               return $img                                          
                                             else 
                                                 ""

                                let $img-meta-title := if($pic eq "") then () else
                                                         let $cust-uri := xdmp:document-properties($uri)//ppt:shapetags/ppt:shape/ppt:tag[@ppt:rid=$tag-rid]/ppt:custompart
                                                         return fn:data(fn:doc($cust-uri)//dc:title)


                                             (: this is a mess, so take just start of text :)
				let $snippet := if(fn:not(fn:empty($img-meta-title))) then
                                                    $img-meta-title
                                                else if(fn:string-length(fn:data($ctrl)) > 120) then 
                                                   fn:concat(fn:substring(fn:data($ctrl), 1, 120), "...") 
                                                else fn:data($ctrl)


                                let $icon-type := if (fn:local-name($hit/p:pic) eq "pic") then
                                                      "imageIcon"
                                                  else if(fn:local-name($hit/p:sp) eq "sp") then
                                                      "textIcon"
                                                  else "slideIcon"

				return 
                                 <div class="searchreturnresult">
                                  <h4> 
                                      <a href="./utils/openpkg.xqy?uri={xdmp:url-encode($uri)}" onmouseup="blurSelected(this)" class="blacklink">
                                          {fn:data($hit/@title)}  <!-- presentation name  -->
                                      </a>
                                  </h4> 
                                 <p class="byline">{fn:concat("Modified: ",fn:data($hit/@moddate))}&nbsp; 
                                          <span>{fn:data($hit/@modby)}</span> <!-- presentation metadata  -->
                                 </p>

<!-- if preso, list of presos only, make title link to open, if opened, it already has tags/metadata -->
<!-- if slide, use existing sample way of inserting slide, though we need to return tags/metadata -->
<!-- if component , then need to send json as well as metadata -->
				     { let $page := if($search-type eq "presentation") then
                                                     let $new-uri :=fn:concat(fn:substring-before($slideuri,"_PNG"),"_PNG/Slide1.PNG")
                                                     let $new-src := fn:concat("search/download-support.xqy?uid=",$new-uri)
                                                     return  (<img src="{$new-src}" class="resize"></img>,<br/>)
                                                     
                                                    else
                                                      let $highlight := cts:highlight(<p class="searchreturnsnippet" title="{fn:data($ctrl)}">
                                                                                      <span class="{$icon-type}">&nbsp;</span>
                                                                                       {$snippet} </p>, 
                                                                                       $or-query, 
                                                                                      <strong class="ML-highlight">{$cts:text}</strong>)  
                                                      let $img := <img src="{$src}" class="resize"></img>
                                                      let $action :=  <div id="searchresultactions">
					    <!--<a href="#" class="insertbtn" OnClick="InsertAction('{xdmp:url-encode($uri)}', '{$path}');">INSERT</a>&nbsp;-->
                                            {
(:was passing $path, now passing $tag-rid, which may be empty for slides with no tags :) 
                                              if ($search-type eq "component") then
					        <a href="javascript:insertComponentAction('{xdmp:url-encode($uri)}', '{$tag-rid}','{$pic}','{$idx}');" onmouseup="blurSelected(this)" class="smallbtn searchinsertbtn"><span>Insert</span></a>
                                              else 
                                                   <a href="javascript:insertSlideAction('{xdmp:url-encode($uri)}', '{$tag-rid}','{xdmp:url-encode($pptxuri)}','{$slide-idx}','{$idx}');" onmouseup="blurSelected(this)" class="smallbtn searchinsertbtn"><span>Insert</span></a>
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

 
