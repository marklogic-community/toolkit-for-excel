xquery version "1.0-ml";
(: Copyright 2009-2010 Mark Logic Corporation

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

module namespace  ppt = "http://marklogic.com/openxml/powerpoint";

declare namespace a="http://schemas.openxmlformats.org/drawingml/2006/main";
declare namespace w="http://schemas.openxmlformats.org/wordprocessingml/2006/main";
declare namespace v="urn:schemas-microsoft-com:vml";
declare namespace ve="http://schemas.openxmlformats.org/markup-compatibility/2006";
declare namespace o="urn:schemas-microsoft-com:office:office";
declare namespace r="http://schemas.openxmlformats.org/officeDocument/2006/relationships";
declare namespace rel= "http://schemas.openxmlformats.org/package/2006/relationships";
declare namespace m="http://schemas.openxmlformats.org/officeDocument/2006/math";
declare namespace wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
declare namespace w10="urn:schemas-microsoft-com:office:word";
declare namespace wne="http://schemas.microsoft.com/office/word/2006/wordml";
declare namespace pkg="http://schemas.microsoft.com/office/2006/xmlPackage";
declare namespace pic="http://schemas.openxmlformats.org/drawingml/2006/picture";
declare namespace pr="http://schemas.openxmlformats.org/package/2006/relationships";
declare namespace types="http://schemas.openxmlformats.org/package/2006/content-types";
declare namespace zip="xdmp:zip";


declare default element namespace "http://schemas.openxmlformats.org/package/2006/content-types";

(:version 1.0-3:)

declare function  ppt:ct-utils-update-types(
   $content-types as node(), 
   $start-idx as xs:integer,
   $img-types as xs:string*, 
   $theme-ids as xs:string*
) as element(types:Types)
{

    let $ctypes-no-theme := ppt:ct-utils-remove-themes($content-types,$theme-ids)
    let $ctypes-no-hm := ppt:ct-utils-remove-hm($ctypes-no-theme)
    let $upd-ctypes := ppt:ct-utils-add-slide($ctypes-no-hm,$start-idx)
    return if(fn:empty($img-types)) then 
              $upd-ctypes 
           else
              ppt:ct-utils-add-img-defaults($upd-ctypes,$img-types)
  
};

declare function ppt:ct-utils-remove-themes(
   $content-types as node(),
   $theme-ids as xs:string*
) as element(types:Types)
{
     let $children := $content-types/node()
     let $themes  := $content-types/Override[fn:matches(@PartName,"theme\d+\.xml?")] (:add if below here:)
     let $override := $content-types/Override[fn:not(fn:matches(@PartName,"theme\d+\.xml?"))]
     let $default := $content-types/Default
   
     let $upd-themes := for $t in $themes
                         return if(fn:substring-before(fn:substring-after($t/@PartName,"/ppt/theme/theme"),".xml") = $theme-ids) then () else $t
     
    
     return element{fn:name($content-types)} {$content-types/@*, $upd-themes, $override, $default} 
};

declare function ppt:ct-utils-add-slide(
   $content-types as node(), 
   $slide-idx as xs:integer
) as element(types:Types)
{
	let $slidename := fn:concat("/ppt/slides/slide",$slide-idx,".xml")
	let $overrideelem := element Override {attribute PartName{$slidename }, attribute ContentType {"application/vnd.openxmlformats-officedocument.presentationml.slide+xml" } }
	let $test := $content-types/Override[fn:ends-with(@PartName,$slidename)]
	let $final := if(fn:empty($test)) then
                 	let $children := $content-types/node() 
                 	return element{fn:name($content-types)} {$content-types/@*, $children, $overrideelem}
                      else  
                  (:some other function adjust all slides, add this one, blah :)
                   (:add function to test for image types and add :)
                    (:let $pngDefTest := <Default Extension="png" ContentType="image/png"/> :)
                  
                    	let $non-slide-types := $content-types/Override[fn:not(fn:matches(@PartName,"/ppt/slides/slide\d+\.xml"))] 
                    	let $defaults := $content-types/Default
                    	let $slide-types :=$content-types/Override[fn:matches(@PartName,"/ppt/slides/slide\d+\.xml")] 
                    	let $upd-slide-types := 
                                            for $s in $slide-types
                                            let $o-slideIdx := xs:integer(fn:substring-before(fn:substring-after($s/@PartName, "slides/slide"),".xml"))
                                            let $finSld := if($o-slideIdx >= $slide-idx) 
                                                           then 
                                                              let $new-slidename := fn:concat("/ppt/slides/slide",($o-slideIdx+1),".xml")
                                                              return element Override {attribute PartName{$new-slidename }, attribute ContentType {"application/vnd.openxmlformats-officedocument.presentationml.slide+xml" } } 
                                                           else
                                                             $s
                                            return $finSld
                   	return  element{fn:name($content-types)} {$content-types/@*, $defaults, $non-slide-types,$upd-slide-types, $overrideelem} 

	return $final 

};

declare function ppt:ct-utils-remove-hm( 
   $content-types as node()
) as element(types:Types)
{
 (:path expression:)
   	let $children := $content-types/node()
   	let $finalchildren := 
                         for $c in $children
                         let $n := if(fn:matches($c/@PartName,"handoutMaster\d+\.xml")) then () else $c
                         return $n
                          
   	return element{fn:name($content-types)} {$content-types/@*, $finalchildren}
};

(: CHANGE add image defaults :)
declare function  ppt:ct-utils-add-img-defaults(
   $content-types as node(), 
   $img-types as xs:string*
) as element(types:Types)
{
(: loop thru default,  :)
(: $ctypes/Default[not(@Extension = $types :)

        let $new-types := for $t in $img-types
                          let $ext := $t
                          let $ct := fn:concat("image/",$t)
                          return element Default {attribute Extension{$ext}, attribute ContentType{$ct}}
                        
        let $default  := $content-types/Default
        let $all-def := ($new-types,$default)
   
        let $dist := fn:distinct-values($all-def/@Extension)
        let $final-def := for $d in $dist
                          return $all-def[@Extension = $d][1]

        
        let $other  := $content-types/* except $content-types/Default

        
        return element{fn:name($content-types)} {$content-types/@*,($final-def,$other)}
};

