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

(: ====== BEGIN remove themes from content-types :)
(: gotta be a better way to do this , to much looping:)
declare function ppt:ct-remove-theme($ctypes as node(), $theme-ids as xs:string*) as node()*
{
  for $t in $theme-ids
  let $override := if(fn:ends-with($ctypes/@PartName, $t)) then () else $ctypes 
  return $override

};
declare function ppt:passthru-ct-remove-theme($x as node(), $theme-ids as xs:string*) as node()*
{
   for $i in $x/node() return ppt:dispatch-ct-remove-theme($i, $theme-ids)
};

declare function ppt:dispatch-ct-remove-theme($ctypes as node(), $theme-ids as xs:string*) as node()*
{
      typeswitch($ctypes)
       case text() return $ctypes
       case document-node() return   document{ppt:passthru-ct-remove-theme($ctypes, $theme-ids)} 
       case element(types:Override) return ppt:ct-remove-theme($ctypes, $theme-ids) 
       case element() return  element{fn:name($ctypes)} {$ctypes/@*,passthru-ct-remove-theme($ctypes, $theme-ids)}
       default return $ctypes

};

declare function  ppt:ct-utils-remove-theme($ctypes as node(), $theme-ids as xs:string*)
{
  ppt:dispatch-ct-remove-theme($ctypes, $theme-ids)
};

(: ====== END remove themes from content-types :)

(: ====== BEGIN add slide  to content-types :)
declare function ppt:ct-utils-add-slide($ctypes as node(), $slide-idx as xs:integer)
{
let $slidename := fn:concat("/ppt/slides/slide",$slide-idx,".xml")
let $overrideelem := element Override {attribute PartName{$slidename }, attribute ContentType {"application/vnd.openxmlformats-officedocument.presentationml.slide+xml" } }
let $test := $ctypes/Types/Override[fn:ends-with(@PartName,$slidename)]
let $final := if(fn:empty($test)) then
                 let $children := $ctypes/node() 
                 return element{fn:name($ctypes)} {$ctypes/@*, $children, $overrideelem}
              else (:some other function adjust all slides, add this one, blah :) 
                 ()

return $final (: ($overrideelem, $test, $ctypes) :)

(: <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/> :)
  (: ($slide-idx, $ctypes) :)
};
(: ====== END add slide to content-types :)

(: ====== BEGIN remove handoutMasters from content-types :)
declare function ppt:ct-utils-remove-hm($ctypes as node())
{
   let $children := $ctypes/node()
   let $finalchildren := for $c in $children
                         let $n := if(fn:matches($c/@PartName,"handoutMaster\d+\.xml")) then () else $c
                         return $n
                          
   return element{fn:name($ctypes)} {$ctypes/@*, $finalchildren}
};
(: ====== END  remove handoutMasters from content-types :)






(: ====== BEGIN remove themes from content-types :)
(: ====== END remove themes from content-types :)
(: ====== BEGIN remove themes from content-types :)
(: ====== END remove themes from content-types :)
(: ====== BEGIN remove themes from content-types :)
(: ====== END remove themes from content-types :)
(: ====== BEGIN remove themes from content-types :)
(: ====== END remove themes from content-types :)
