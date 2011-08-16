xquery version "1.0-ml";

declare namespace html = "http://www.w3.org/1999/xhtml";
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

import module namespace rest = "http://marklogic.com/appservices/rest"
        at "/MarkLogic/appservices/utils/rest.xqy";

import module namespace json = "http://marklogic.com/json" 
          at "/MarkLogic/appservices/utils/json.xqy";

import module namespace requests =   "http://marklogic.com/appservices/requests" at "requests.xqy";

declare variable $LOOKAHEAD-PAGES as xs:integer := 1;
declare variable $SIZE := 10;

declare function ps:page-results(
 $query as cts:query,
 $start as xs:integer)
 as element(ps:results-page)
{
  let $page-stop := $start + $SIZE -1
  let $stop := 1 + $page-stop + ($SIZE * $LOOKAHEAD-PAGES)
  let $results := cts:search(//p:sld, $query)[ $start to $stop ]
                              
  return
    (: if we stepped off of the end, recurse to the previous page :)
    if(fn:not(fn:empty($results)) and ($start - $SIZE) gt 1)
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
          
          let $node-uri := xdmp:node-uri($r)
  
          let $single := fn:data(xdmp:document-properties($node-uri)/prop:properties/ppt:single)

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
           attribute single { $single },
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
           attribute slide-index { $slide-idx },
           $r
          }
      
      }

};


declare function ps:get-and-query($raw as xs:string)
 as cts:query?
{ 
  let $words := tokenize($raw, '\s+')[. ne '']
  where $words
  return cts:and-query($words)
};



let $request := $requests:options/rest:request[@endpoint = "/search.xqy"][1]

let $map  := rest:process-request($request) 

let $q := xdmp:url-decode(map:get($map, "q"))
let $format := if(fn:empty(map:get($map, "format"))) then 
                  "xml" 
               else 
                  map:get($map,"format")

let $start := if(fn:empty(map:get($map, "start"))) then 
                  0 
              else
                  map:get($map,"start")

(:let $end := $start + 9 :)

let $and-query := ps:get-and-query($q)
let $results := ps:page-results($and-query,$start)
let $remainder := $results/@remainder
let $r-start := $results/@start
let $estimate := $results/@estimated

let $package :=
 <results>
<remainder>{fn:data($remainder)}</remainder>
<start>{fn:data($r-start)}</start>
<estimated>{fn:data($estimate)}</estimated>

{
for $r in $results/ps:result
let $pptx-uri := fn:data($r/@docuri)
let $slide-index := fn:data($r/@slide-index)
let $slide-img := fn:concat($pptx-uri,"/slides/slide",$slide-index)
return <result>
          <slide>
            <image>{$slide-img}</image>
            <index>{$slide-index}</index>
            <single>{ fn:data($r/@single)}</single>
            <uri>{fn:data($r/@uri)}</uri>
          </slide>
          <pptx-uri>{$pptx-uri}</pptx-uri>
          <pptx-title>{fn:data($r/@title)}</pptx-title>
          <lastmodby>{fn:data($r/@modby)}</lastmodby>
          <lastmoddate>{fn:data($r/@moddate)}</lastmoddate>
       </result>
}</results>
return if($format eq "json") then
          json:serialize($package)
       else
         $package

                  
