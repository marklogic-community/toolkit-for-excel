xquery version "1.0-ml";
declare namespace html = "http://www.w3.org/1999/xhtml";
let $server := xdmp:get-request-field("srvuri")

 (:"http://localhost:8060/office/presentations":)
let $slideuri := xdmp:get-request-field("slduri")
let $dest := xdmp:get-request-field("dest")

let $results := if($dest eq "workspace") then 

(:"/paven/one.pptx/slides":)
                  let $slides := xdmp:http-get(fn:concat($server,$slideuri))[2]/node()
                  let $imageuris := for $s in $slides/slide
                                    return fn:concat($server, fn:string($s/image))
                  return
                    xdmp:quote(<ul class="connect">
                                {
                                 for $img in $imageuris
                                 let $src := $img
                                 (:let $src := fn:concat("download-support.xqy?uri=",$img):)
                                 return <li><img src="{$src}"/></li>
                                }
                               </ul>)
                else 
                  let $slides := xdmp:http-get($slideuri)[2]/node()
                  let $imageuris := for $i in $slides/slides/slide
                                    return fn:concat($server, fn:string($i/image))
                  return
                    xdmp:quote(<ul class="connect">
                                {
                                 for $img in $imageuris
                                 let $src := $img
                                 return <li><img src="{$src}"/></li>
                                }
                               </ul>)            
return $results

