xquery version "1.0-ml";
declare namespace xladd="http://marklogic.com/openxl/exceladdin";
declare variable $xladd:bsv as xs:string external;

(: define variable $searchparam as xs:string external :)
let $searchval := if(fn:empty($xladd:bsv) or $xladd:bsv eq "") then () else $xladd:bsv
let $header:=((: xdmp:set-response-content-type('text/html'), :)
              <div id="header">
                 <form id="basicsearch" action="insert-image.xqy" method="post">
                   <div>
                      <input type="text" size="40" name="xladd:bsv" autocomplete="off" value={$searchval} id="bsearchval"  method="post"/>&nbsp;
                     <!-- TEST : { $no:color}--><input type="submit" value="Search"/> 
                       
                   </div> 
                  </form>    
             </div>)

return $header
