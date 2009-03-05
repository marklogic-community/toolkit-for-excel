xquery version "1.0-ml";
declare namespace xladd="http://marklogic.com/openxl/exceladdin";
declare namespace ms = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
declare namespace html = "http://www.w3.org/1999/xhtml";
declare variable $xladd:bsv as xs:string external;

declare function xladd:get-content-parent($uri as xs:string, $param as xs:string)
{
  let $doc := fn:doc($uri)
  let $hlt := cts:highlight($doc,$param, <marker>{$cts:text}</marker>)
  let $mark := fn:node-name($hlt//marker/parent::*/parent::*/parent::*)
  return $mark
};

(: define variable $searchparam as xs:string external :)
let $searchparam := $xladd:bsv
let $space-test := fn:tokenize($searchparam," ")

let $word :=  cts:word-query($searchparam)
let $element := if(fn:count($space-test) gt 1) then
                    ()
                else cts:element-query(xs:QName($searchparam),cts:and-query(()))

let $docs := cts:search(collection(), cts:or-query(($word,$element)))
(: let $elem-test :=  xdmp:node-kind(($docs[1]//$searchparam[1])) :)
let $results
         := for $d in $docs
            let $uri := xdmp:node-uri($d)
            let $tmpclean := if(fn:contains($uri,"parts")) then
                               let $s := fn:substring-before($uri,"_parts")
                               let $s2 := fn:replace($s,"_",".")
                               return $s2
                             else $uri
            return $tmpclean

let $distinct := fn:distinct-values($results)

let $html := for $d in $distinct
             let $div := if(fn:ends-with($d,".xlsx")) then
                         <div id="results">
                            <a href="#" onclick="testOpen('{$d}')">{$d}</a>
                         </div>
                         else
                              let $finaldiv := 
                                           if(fn:count($space-test) gt 1 ) then
                                             let $parent := xladd:get-content-parent($d, $searchparam) 
                                             let $return := 
                                              if(fn:empty($parent)) then () 
                                              else
                                               <div id="results">
                                                <a href="tabletoexcelgenerate.xqy?elemname={$parent}&amp;docuri={$d}">{$d} </a>
                                               </div>
                                             return $return
                                           else
                                             let $parent := xladd:get-content-parent($d, $searchparam)
                                             let $return := 
                                              if(fn:empty($parent)) then 
                                               <div id="results">
                                                <a href="tabletoexcelgenerate.xqy?elemname={$searchparam}&amp;docuri={$d}">{$d} </a>
                                               </div>
                                              else
                                               <div id="results">
                                                <a href="tabletoexcelgenerate.xqy?elemname={$parent}&amp;docuri={$d}">{$d} </a>
                                               </div>
                                           return $return
                              return $finaldiv

             return $div
return if(fn:count($html) eq 0) then <div id="ML-Message"><p><strong>No results returned.</strong></p><p>Your search for "{$searchparam}" did not match anything.</p></div> else $html
