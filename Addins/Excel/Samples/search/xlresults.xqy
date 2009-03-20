xquery version "1.0-ml";
declare namespace xladd="http://marklogic.com/openxl/exceladdin";
declare namespace ms = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
declare namespace html = "http://www.w3.org/1999/xhtml";
declare namespace dc = "http://purl.org/dc/elements/1.1/";
declare namespace cp = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
declare namespace dcterms="http://purl.org/dc/terms/";
declare namespace excel = "http://marklogic.com/openxml/excel";
import module "http://marklogic.com/openxml/excel" at "/MarkLogic/openxml/excel-ml-support.xqy";


declare variable $xladd:bsv as xs:string external;

declare function xladd:get-content-parent($uri as xs:string, $param as xs:string)
{
    try
    {
        let $doc := fn:doc($uri)
        let $hlt := cts:highlight($doc,$param, <marker>{$cts:text}</marker>)
        let $mark := fn:node-name($hlt//marker/parent::*/parent::*/parent::*)
        return $mark
    }catch($exception)
    {
          ()
    }
};

let $searchparam := $xladd:bsv
let $space-test := fn:tokenize($searchparam," ")

let $word :=  cts:word-query($searchparam)
let $element := if(fn:count($space-test) gt 1) then
                    ()
                else cts:element-query(xs:QName($searchparam),cts:and-query(()))

let $docs := cts:search(collection(), cts:or-query(($word,$element)))
let $results
         := for $d in $docs
            let $uri := xdmp:node-uri($d)
            let $tmpclean := if(fn:contains($uri,"parts")) then
                               let $s := fn:substring-before($uri,"_parts")
                               let $s2 := fn:replace($s,"_",".")
                               return $s2
                             else $uri
            return $tmpclean

(: want distinct values, searchparam could be found w/in mulitple worksheets within single workbook :)
let $distinct := fn:distinct-values($results)

let $html := for $d in $distinct
             let $div := if(fn:ends-with($d,".xlsx")) then
                         <div id="results">
                         {
                           (:get directory for parts so we can look at metadata/other :)
                            let $docfolder := fn:replace($d,".xlsx","_xlsx")
                            let $props := fn:concat($docfolder,"_parts/docProps/core.xml")

                           (:metadata:) 
                            let $propsdoc := fn:doc($props)
                            let $lastmodby := if(fn:empty($propsdoc//cp:lastModifiedBy//text())) then () 
                                              else fn:concat("lastmodifiedby: ",$propsdoc//cp:lastModifiedBy//text())
                                              
                                              
                            let $lastmoddate := if(fn:empty($propsdoc//dcterms:modified//text())) then ()
                                                else fn:concat("lastmodified: ",$propsdoc//dcterms:modified//text())

                           (:count:)
                            let $sheetsdir := fn:concat($docfolder,"_parts/xl/worksheets/")
                            let $sheets := excel:directory-uris($sheetsdir)
                            let $count := for $s in $sheets
                                          let $doc := fn:doc($s)
                                          let $hlt := cts:highlight($doc,$searchparam, <marker>{$cts:text}</marker>)
                                          return fn:count($hlt//marker)
                            let $sum := fn:sum($count)
                             
                           (:need to work on html:)                           
                            let $link :=    <table border="0">
                                             <tr>
                                              <td>{$sum}</td>
                                              <td>&nbsp;&nbsp;</td>
                                              <td><a href="#" onclick="openXlsx('{$d}')">{$d}</a></td>
                                             </tr>
                                             </table>
                            let $metadata := <table border="0">
                                             <tr>
                                              <td><img src="excel_icon.bmp"/></td>
                                              <td>&nbsp;&nbsp;</td>
                                              <td valign="middle">{$lastmodby}</td>
                                              <td>&nbsp;&nbsp;</td>
                                              <td valign="middle">{$lastmoddate}</td>
                                             </tr>
                                             </table>
                            return ($link,$metadata)
                         }
                         <br/>
                         </div>
                         else 
                              let $finaldiv := 
                                           if(fn:count($space-test) gt 1 ) then ()
                                           else
                                            (:count:)
                                            let $doc := fn:doc($d)
                                            let $hlt := cts:highlight($doc,$searchparam, <marker>{$cts:text}</marker>)
                                            let $sum := fn:count($hlt//marker)
                                             
                                            (: if no hits, then search result back as result of element name :)
                                            let $parent := if($sum eq 0) then $searchparam else xladd:get-content-parent($d, $searchparam)
                                            let $return := 
                                              if(fn:empty($parent)) then 
                                               <div id="results">
                                                   <table border="0">
                                                    <tr>
                                                     <td>{$sum}</td>
                                                     <td>&nbsp;&nbsp;</td>
                                                     <td><a href="tabletoexcelgenerate.xqy?elemname={$searchparam}&amp;docuri={$d}">{$d} </a></td>
                                                    </tr>
                                                    </table>
                                                 { (: could do test here to determine if xml is likely to be table :) }
                                               </div>
                                              else
                                               <div id="results">
                                                <table border="0">
                                                    <tr>
                                                     <td>{$sum}</td>
                                                     <td>&nbsp;&nbsp;</td>
                                                     <td><a href="tabletoexcelgenerate.xqy?elemname={$parent}&amp;docuri={$d}">{$d} </a></td>
                                                    </tr>
                                                    </table>
                                               </div>
                                           return ($return,<table border="0"><tr><td><img src="xml_icon.gif"/></td></tr></table>,<br/>)

                              return $finaldiv

             return $div
return if(fn:count($html) eq 0) then <div id="ML-Message"><p><strong>No results returned.</strong></p><p>Your search for "{$searchparam}" did not match anything.</p></div> else $html

