xquery version "1.0-ml";
(: Copyright 2009 Mark Logic Corporation

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


declare function  ppt:ct-utils-update-types($ctypes as node(), $start-idx as xs:integer,$types as xs:string*, $theme-ids as xs:string*)
{

    let $ctypes-no-theme := ppt:ct-utils-remove-themes($ctypes,$theme-ids)
    let $ctypes-no-hm := ppt:ct-utils-remove-hm($ctypes)
    let $upd-ctypes := ppt:ct-utils-add-slide($ctypes-no-hm,$start-idx)
    let $final-ctypes := if(fn:empty($types)) then 
                             $upd-ctypes 
                         else
                             ppt:ct-utils-add-defaults($upd-ctypes,$types)
    
    return $final-ctypes
};

declare function ppt:ct-utils-remove-themes($ctypes as node(),$theme-ids as xs:string*)
{
     let $children := $ctypes/node()
     let $themes  := $ctypes/Override[fn:matches(@PartName,"theme\d+\.xml?")]
     let $override := $ctypes/Override[fn:not(fn:matches(@PartName,"theme\d+\.xml?"))]
     let $default := $ctypes/Default
   
     let $upd-themes := for $t in $themes
                         return if(fn:substring-after($t/@PartName,"/ppt/theme/") = $theme-ids) then () else $t
     
    
     return element{fn:name($ctypes)} {$ctypes/@*, $upd-themes, $override, $default}
};

declare function ppt:ct-utils-add-slide($ctypes as node(), $slide-idx as xs:integer)
{
	let $slidename := fn:concat("/ppt/slides/slide",$slide-idx,".xml")
	let $overrideelem := element Override {attribute PartName{$slidename }, attribute ContentType {"application/vnd.openxmlformats-officedocument.presentationml.slide+xml" } }
	let $test := $ctypes/Override[fn:ends-with(@PartName,$slidename)]
	let $final := if(fn:empty($test)) then
                 	let $children := $ctypes/node() 
                 	return element{fn:name($ctypes)} {$ctypes/@*, $children, $overrideelem}
                      else  
                  (:some other function adjust all slides, add this one, blah :)
                   (:add function to test for image types and add :)
                    (:let $pngDefTest := <Default Extension="png" ContentType="image/png"/> :)
                  
                    	let $non-slide-types := $ctypes/Override[fn:not(fn:matches(@PartName,"/ppt/slides/slide\d+\.xml"))] 
                    	let $defaults := $ctypes/Default
                    	let $slide-types :=$ctypes/Override[fn:matches(@PartName,"/ppt/slides/slide\d+\.xml")] 
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
                   	return  element{fn:name($ctypes)} {$ctypes/@*, $defaults, $non-slide-types,$upd-slide-types, $overrideelem} 

	return $final 

};

(: ====== BEGIN remove handoutMasters from content-types :)
declare function ppt:ct-utils-remove-hm($ctypes as node())
{
   	let $children := $ctypes/node()
   	let $finalchildren := 
                         for $c in $children
                         let $n := if(fn:matches($c/@PartName,"handoutMaster\d+\.xml")) then () else $c
                         return $n
                          
   	return element{fn:name($ctypes)} {$ctypes/@*, $finalchildren}
};
(: ====== END  remove handoutMasters from content-types :)
(:CHANGE add image defaults :)
declare function  ppt:ct-utils-add-defaults($ctypes as node(), $types as xs:string*)
{
        let $new-types := for $t in $types
                          let $ext := $t
                          let $ct := fn:concat("image/",$t)
                          return element Default {attribute Extension{$ext}, attribute ContentType{$ct}}
                        
        let $default  := $ctypes/Default
        let $all-def := ($new-types,$default)
   
        let $dist := fn:distinct-values($all-def/@Extension)
        let $final-def := for $d in $dist
                          return $all-def[@Extension = $d][1]

        
        let $other  := $ctypes/* except $ctypes/Default

        
        return element{fn:name($ctypes)} {$ctypes/@*,($final-def,$other)}
};


(: ====== BEGIN remove themes from content-types :)
(: ====== END remove themes from content-types :)
(: ====== BEGIN remove themes from content-types :)
(: ====== END remove themes from content-types :)
(: ====== BEGIN remove themes from content-types :)
(: ====== END remove themes from content-types :)
(: ====== BEGIN remove themes from content-types :)
(: ====== END remove themes from content-types :)
