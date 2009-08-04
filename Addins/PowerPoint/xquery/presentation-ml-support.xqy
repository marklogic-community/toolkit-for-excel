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
declare namespace rel="http://schemas.openxmlformats.org/package/2006/relationships";
declare namespace m="http://schemas.openxmlformats.org/officeDocument/2006/math";
declare namespace wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
declare namespace w10="urn:schemas-microsoft-com:office:word";
declare namespace wne="http://schemas.microsoft.com/office/word/2006/wordml";
declare namespace pkg="http://schemas.microsoft.com/office/2006/xmlPackage";
declare namespace pic="http://schemas.openxmlformats.org/drawingml/2006/picture";
declare namespace pr="http://schemas.openxmlformats.org/package/2006/relationships";
declare namespace types="http://schemas.openxmlformats.org/package/2006/content-types";
declare namespace zip="xdmp:zip";
declare namespace p="http://schemas.openxmlformats.org/presentationml/2006/main";

import module "http://marklogic.com/openxml/powerpoint" at "/MarkLogic/openxml/presentation-ml-support-content-types.xqy"; 

declare default element namespace "http://schemas.openxmlformats.org/package/2006/relationships";


(: ================================== :)

declare function ppt:formatbinary($s as xs:string*) as xs:string*
{

  (: xdmp:invoke("formatbinary.xqy",( xs:QName("ppt:image"), $s)) :)

 if(fn:string-length($s) > 0) then
     let $firstpart := fn:concat(fn:substring($s,1,76))
      let $tail := fn:substring-after($s,$firstpart) 
      let $tail := fn:substring($s,77) 
     return ($firstpart,ppt:formatbinary($tail))
                  else
             ()

};

declare function ppt:get-part-content-type($uri as xs:string) as xs:string?
{
   if(fn:ends-with($uri,".rels"))
   then 
        "application/vnd.openxmlformats-package.relationships+xml"
   else if(fn:ends-with($uri,"glossary/document.xml"))
   then
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml"
   else if(fn:ends-with($uri,"presentation.xml"))
   then
      "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml" 
   else if(fn:matches($uri, "slide\d+\.xml"))
   then 
      "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
   else if(fn:matches($uri, "notesSlide\d+\.xml"))
   then 
      "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"
   else if(fn:matches($uri, "slideMaster\d+\.xml"))
   then 
      "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"
   else if(fn:matches($uri, "slideLayout\d+\.xml"))
   then
      "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"
   else if(fn:matches($uri,"theme\d+\.xml"))
   then
       "application/vnd.openxmlformats-officedocument.theme+xml"
   else if(fn:matches($uri,"notesMaster\d+\.xml"))
   then
       "application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml"
   else if(fn:matches($uri,"handoutMaster\d+\.xml"))
   then
       "application/vnd.openxmlformats-officedocument.presentationml.handoutMaster+xml"
   else if(fn:ends-with($uri,"presProps.xml"))
   then
       "application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"
   else if(fn:ends-with($uri,"viewProps.xml"))
   then
       "application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"
   
   else if(fn:ends-with($uri,"tableStyles.xml"))
   then
       "application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"


   else if(fn:ends-with($uri,"styles.xml"))
   then 
      "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"
   else if(fn:ends-with($uri,"webSettings.xml"))
   then 
      "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml"
   (: else if(fn:ends-with($uri,"word/fontTable.xml")) :)
   else if(fn:ends-with($uri,"fontTable.xml"))
   then 
      "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"
   else if(fn:ends-with($uri,"word/footnotes.xml"))
   then 
      "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"
   else if(fn:matches($uri, "header\d+\.xml"))
   then 
      "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
   else if(fn:matches($uri, "footer\d+\.xml"))
   then 
      "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"
   else if(fn:ends-with($uri,"word/endnotes.xml"))
   then
      "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"
   else if(fn:ends-with($uri,"docProps/core.xml"))
   then
      "application/vnd.openxmlformats-package.core-properties+xml"
   else if(fn:ends-with($uri,"docProps/app.xml"))
   then
      "application/vnd.openxmlformats-officedocument.extended-properties+xml"
   else if(fn:ends-with($uri,"docProps/custom.xml")) 
   then
      "application/vnd.openxmlformats-officedocument.custom-properties+xml"
   else if(fn:ends-with($uri,"jpeg")) 
   then
      "image/jpeg"
   else if(fn:ends-with($uri,"wmf")) 
   then
      "image/x-wmf"
   else if(fn:ends-with($uri,"png")) 
   then
      "image/png"
   else if(fn:ends-with($uri,"gif"))
   then
       "image/gif"
   else if(fn:matches($uri,"customXml/itemProps\d+\.xml")) then
      "application/vnd.openxmlformats-officedocument.customXmlProperties+xml"
   else if(fn:matches($uri,"customXml/item\d+\.xml")) then
      "application/xml"
   else
       ()
    
};

declare function ppt:get-part-attributes($uri as xs:string) as node()*
{
  let $cleanuri := fn:replace($uri,"\\","/")
  let $name := attribute pkg:name{$cleanuri}
  let $contenttype := attribute pkg:contentType{ppt:get-part-content-type($cleanuri)}
  let $padding := if(fn:ends-with($cleanuri,".rels")) then

                     if(fn:starts-with($cleanuri,"/ppt/glossary") or 
                        fn:starts-with($cleanuri,"/ppt/slides/_rels") or
                        fn:starts-with($cleanuri,"/ppt/notesSlides/_rels") or
                        fn:starts-with($cleanuri,"/ppt/slideLayouts/_rels") or
                        fn:starts-with($cleanuri,"/ppt/slideMasters/_rels")
                       ) then
                         ()
                    
                     else if(fn:starts-with($cleanuri,"/_rels")) then
                      attribute pkg:padding{ "512" }
                     else    
                      attribute pkg:padding{ "256" }
                  else
                     ()
  let $compression := if(fn:ends-with($cleanuri,"jpeg") or fn:ends-with($cleanuri,"png")) then 
                         attribute pkg:compression { "store" } 
                      else ()
  
  return ($name, $contenttype, $padding, $compression)
};

declare function ppt:get-package-part($directory as xs:string, $uri as xs:string) as node()?
{
  let $fulluri := $uri
  let $docuri := fn:concat("/",fn:substring-after($fulluri,$directory))
  let $data := fn:doc($fulluri)

  let $part := if(fn:empty($data) or fn:ends-with($fulluri,"[Content_Types].xml")) then () 
               else if(fn:ends-with($fulluri,".jpeg") or fn:ends-with($fulluri,".wmf") or fn:ends-with($fulluri,".png")) then
                  let $bin :=   xs:base64Binary(xs:hexBinary($data)) cast as xs:string 
                    let $formattedbin := fn:string-join(ppt:formatbinary($bin),"&#x9;&#xA;") 
                  return  element pkg:part { ppt:get-part-attributes($docuri), element pkg:binaryData { $formattedbin  }   }
               else
                  element pkg:part { ppt:get-part-attributes($docuri), element pkg:xmlData { $data }}
  return  $part 
};

declare function ppt:make-package($directory as xs:string, $uris as xs:string*) as node()*
{
  let $package := element pkg:package { 
                            for $uri in $uris
                            let $part := ppt:get-package-part($directory,$uri)
                            return $part }
                           
return $package
(: <?mso-application progid="Word.Document"?>, $package :)
(: <?mso-application progid="PowerPoint.Show"?> :)
};

declare function ppt:package-uris-from-directory($docuri as xs:string) as xs:string*
{

  cts:uris("","document",cts:directory-query($docuri,"infinity"))

};

declare function ppt:package-uris-from-directory($docuri as xs:string, $depth as xs:string) as xs:string*
{

  cts:uris("","document",cts:directory-query($docuri,$depth))

};

declare function ppt:package-files-only($uris as xs:string*) as xs:string*
{
                  for $uri in $uris
                  let $u := if(fn:ends-with($uri,"/")) then () else $uri
                  return $u
};

(: ===================== BEGIN file and dir helpers ====================== :)
declare function ppt:uri-content-types($dir as xs:string?) as xs:string
{
  fn:concat($dir,"[Content_Types].xml")
};
declare function ppt:uri-rels-dir($dir as xs:string?) as xs:string
{
    fn:concat($dir,"_rels/")
};

declare function ppt:uri-rels($dir as xs:string?) as xs:string
{
    fn:concat(ppt:uri-rels-dir($dir),".rels") 
};

declare function ppt:uri-docprops-dir($dir as xs:string?) as xs:string
{
    fn:concat($dir,"docProps/")
};

declare function ppt:uri-app-props($dir as xs:string?) as xs:string
{
      fn:concat(ppt:uri-docprops-dir($dir),"app.xml")
};

declare function ppt:uri-core-props($dir as xs:string?) as xs:string
{
      fn:concat(ppt:uri-docprops-dir($dir),"core.xml")
};

declare function ppt:uri-ppt-dir($dir as xs:string?) as xs:string
{
      fn:concat($dir, "ppt/")
};

declare function ppt:uri-ppt-rels-dir($dir as xs:string?) as xs:string
{
      fn:concat(ppt:uri-ppt-dir($dir),"_rels/")
};

declare function ppt:uri-ppt-rels($dir as xs:string?) as xs:string
{
     fn:concat(ppt:uri-ppt-rels-dir($dir),"presentation.xml.rels")
};

declare function ppt:uri-ppt-handout-masters-dir($dir as xs:string?) as xs:string
{
      fn:concat(ppt:uri-ppt-dir($dir),"handoutMasters/")
};

declare function ppt:uri-ppt-handout-master-rels-dir($dir as xs:string?) as xs:string
{
      fn:concat(ppt:uri-ppt-handout-masters-dir($dir),"_rels/")
};

declare function ppt:uri-ppt-handout-master($dir as xs:string?, $idx as xs:integer) as xs:string
{
    let $handoutMasterFile := fn:concat("handoutMaster",$idx,".xml")
    return fn:concat(ppt:uri-ppt-handout-masters-dir($dir), $handoutMasterFile)
};

declare function ppt:uri-ppt-handout-master-rels($dir as xs:string?, $idx as xs:integer) as xs:string
{
    let $handoutMasterRelsFile := fn:concat("handoutMaster",$idx,".xml.rels")
    return fn:concat( ppt:uri-ppt-handout-master-rels-dir($dir),$handoutMasterRelsFile)
};

declare function ppt:uri-ppt-media-dir($dir as xs:string?) as xs:string
{
     fn:concat(ppt:uri-ppt-dir($dir),"media/")
};

declare function ppt:uri-ppt-notes-masters-dir($dir as xs:string?) as xs:string
{
     fn:concat(ppt:uri-ppt-dir($dir),"notesMasters/")
};

declare function ppt:uri-ppt-notes-masters-rels-dir($dir as xs:string?) as xs:string
{
    fn:concat(ppt:uri-ppt-notes-masters-dir($dir),"_rels/")
}; 

declare function ppt:uri-ppt-notes-master($dir as xs:string?, $idx as xs:integer) as xs:string
{
    let $notesMasterFile := fn:concat("notesMaster",$idx,".xml")
    return fn:concat(ppt:uri-ppt-notes-masters-dir($dir), $notesMasterFile)
};

declare function ppt:uri-ppt-notes-master-rels($dir as xs:string?, $idx as xs:integer) as xs:string
{
    let $notesMasterRelsFile := fn:concat("notesMaster",$idx,".xml.rels")
    return fn:concat( ppt:uri-ppt-notes-masters-rels-dir($dir),$notesMasterRelsFile)
};

declare function ppt:uri-ppt-slide-layouts-dir($dir as xs:string?) as xs:string
{
     fn:concat(ppt:uri-ppt-dir($dir),"slideLayouts/")
};

declare function ppt:uri-ppt-slide-layout-rels-dir($dir as xs:string?) as xs:string
{
    fn:concat(ppt:uri-ppt-slide-layouts-dir($dir),"_rels/")
}; 

declare function ppt:uri-ppt-slide-layout($dir as xs:string?, $idx as xs:integer) as xs:string
{
    let $slideLayoutFile := fn:concat("slideLayout",$idx,".xml")
    return fn:concat(ppt:uri-ppt-slide-layouts-dir($dir), $slideLayoutFile)
};

declare function ppt:uri-ppt-slide-layout-rels($dir as xs:string?, $idx as xs:integer) as xs:string
{
    let $slideLayoutRelsFile := fn:concat("slideLayout",$idx,".xml.rels")
    return fn:concat( ppt:uri-ppt-slide-layout-rels-dir($dir),$slideLayoutRelsFile)
};

declare function ppt:uri-ppt-slide-masters-dir($dir as xs:string?) as xs:string
{
     fn:concat(ppt:uri-ppt-dir($dir),"slideMasters/")
};

declare function ppt:uri-ppt-slide-master-rels-dir($dir as xs:string?) as xs:string
{
    fn:concat(ppt:uri-ppt-slide-masters-dir($dir),"_rels/")
}; 

declare function ppt:uri-ppt-slide-master($dir as xs:string?, $idx as xs:integer) as xs:string
{
    let $slideMasterFile := fn:concat("slideMaster",$idx,".xml")
    return fn:concat(ppt:uri-ppt-slide-masters-dir($dir), $slideMasterFile)
};

declare function ppt:uri-ppt-slide-master-rels($dir as xs:string?, $idx as xs:integer) as xs:string
{
    let $slideMasterRelsFile := fn:concat("slideMaster",$idx,".xml.rels")
    return fn:concat( ppt:uri-ppt-slide-master-rels-dir($dir),$slideMasterRelsFile)
};

declare function ppt:uri-ppt-slides-dir($dir as xs:string?) as xs:string
{
     fn:concat(ppt:uri-ppt-dir($dir),"slides/")
};


declare function ppt:uri-ppt-slide-rels-dir($dir as xs:string?) as xs:string
{
    fn:concat(ppt:uri-ppt-slides-dir($dir),"_rels/")
}; 

declare function ppt:uri-ppt-slide($dir as xs:string?, $idx as xs:integer) as xs:string
{
    let $slideFile := fn:concat("slide",$idx,".xml")
    return fn:concat(ppt:uri-ppt-slides-dir($dir), $slideFile)
};

declare function ppt:uri-ppt-slide-rels($dir as xs:string?, $idx as xs:integer) as xs:string
{
    let $slideRelsFile := fn:concat("slide",$idx,".xml.rels")
    return fn:concat( ppt:uri-ppt-slide-rels-dir($dir),$slideRelsFile)
};

declare function ppt:uri-ppt-theme-dir($dir as xs:string?) as xs:string
{
     fn:concat(ppt:uri-ppt-dir($dir),"theme/")
};

declare function ppt:uri-ppt-theme-rels-dir($dir as xs:string?) as xs:string
{
    fn:concat(ppt:uri-ppt-theme-dir($dir),"_rels/")
}; 

declare function ppt:uri-ppt-theme($dir as xs:string?, $idx as xs:integer) as xs:string
{
    let $themeFile := fn:concat("theme",$idx,".xml")
    return fn:concat(ppt:uri-ppt-theme-dir($dir), $themeFile)
};

declare function ppt:uri-ppt-theme-rels($dir as xs:string?, $idx as xs:integer) as xs:string
{
    let $themeRelsFile := fn:concat("theme",$idx,".xml.rels")
    return fn:concat( ppt:uri-ppt-theme-rels-dir($dir),$themeRelsFile)
};

declare function ppt:uri-ppt-presentation($dir as xs:string?) as xs:string
{
     fn:concat(ppt:uri-ppt-dir($dir),"presentation.xml")
};
declare function ppt:uri-ppt-pres-props($dir as xs:string?) as xs:string
{
     fn:concat(ppt:uri-ppt-dir($dir),"presProps.xml")
};
declare function ppt:uri-ppt-table-styles($dir as xs:string?) as xs:string
{
     fn:concat(ppt:uri-ppt-dir($dir),"tableStyles.xml")
};
declare function ppt:uri-ppt-view-props($dir as xs:string?) as xs:string
{
     fn:concat(ppt:uri-ppt-dir($dir),"viewProps.xml")
};
(: ===================== END file and dir helpers  ====================== :)

declare function ppt:max-file-id($dir as xs:string*, $type as xs:string*, $depth as xs:string)
{
  let $files :=  ppt:package-uris-from-directory($dir,"1")
  let $numbers := if(fn:empty($files)) then 0
                  else
                     for $i in $files
                     let $tmp1 := fn:substring-after($i,fn:concat($dir,$type))
                     let $tmp2 := fn:substring-before($tmp1,".")
                     return xs:integer($tmp2)
  return fn:max($numbers)

};


declare function ppt:max-image-id($dir as xs:string*)
{
  
  ppt:max-file-id($dir, "image", "infinity") 

};

declare function ppt:max-slide-id($dir as xs:string*)
{
  ppt:max-file-id($dir, "slide", "1") 
};

declare function ppt:handout-master-theme-ids($hm as xs:string*) as xs:string*
{
(:need to findout theme#.xml, so we can remove from pkg :)
  let $h-rels := ppt:package-uris-from-directory($hm)
  let $theme-idx := for $rel in $h-rels
                 let $doc := fn:doc($rel)
                 let $theme := $doc/rel:Relationships/rel:Relationship/@Target
                 let $theme-uri := fn:substring-after($theme,"../theme/")
                 return $theme-uri
  return $theme-idx
};

(:may want a generic fileid function, similar to max :)

declare function ppt:image-id($uri as xs:string) as xs:integer
{
  xs:integer(fn:substring-before(fn:substring-after($uri,"image"),"."))
};
(: begin related to updating slide.xml.rels ============================ :)

declare function ppt:update-rels-rel($r as node(), $n-idx as xs:integer)
{
   if(fn:matches($r/@Target,"slideLayout")) then $r 
   else if(fn:matches($r/@Target,"image")) then
    let $target := $r/@Target
    let $prfx := fn:substring-before($target,"image")
    let $sfx := fn:substring-after($target,".")
    let $id := ppt:image-id($target)
    let $n-targ := fn:replace($target, xs:string($id),xs:string($id + $n-idx))
    return  element{fn:name($r)} {$r/@* except $r/@Target, attribute Target{$n-targ}}
   else $r
};
declare function ppt:passthru-rels($x as node(), $idx as xs:integer) as node()*
{
   for $i in $x/node() return ppt:dispatch-slide-rels($i, $idx)
};

declare function ppt:dispatch-slide-rels($rels as node(), $new-img-idx as xs:integer) as node()*
{
      typeswitch($rels)
       case text() return $rels
       (: case document-node() return document{$rels/@*,ppt:passthru-rels($rels, $new-img-idx)} :)
       case document-node() return document{ppt:passthru-rels($rels, $new-img-idx)}
       case element(rel:Relationship) return ppt:update-rels-rel($rels, $new-img-idx) 
       case element(rel:Relationships) return element{fn:name($rels)} {$rels/namespace::*, $rels/@*,passthru-rels($rels, $new-img-idx)}
       case element() return  element{fn:name($rels)} {$rels/@*,passthru-rels($rels, $new-img-idx)}
       default return $rels

};

(:declare function ppt:upd-slide-rels($orig-slide-rels as element(r:Relationships),$img-targs as xs:string*,$new-img-idx as xs:integer) :)
declare function ppt:upd-slide-rels($orig-slide-rels as node(),$img-targs as xs:string*,$new-img-idx as xs:integer)
{
  ppt:dispatch-slide-rels($orig-slide-rels, $new-img-idx)
 (: fn:doc($orig-slide-rels) :)
};


(: end related to updating slide.xml.rels ============================== :)
declare function ppt:slide-and-relationships($t-pres as xs:string, $s-pres as xs:string, $s-idx as xs:integer, $start-idx as xs:integer)
{
(: map needs slide#.xml, 
             slide#.xml.rels (updated accordingly (later?))
             relationships (images, have to getMaxImageId and +1
   may have to grab prior? or just use map later?
   need to potentially update Content-Types with type (gif))
   don't think we need numbers elsewhere, relationship defined in slide.xml.rels
:)

let $smap := map:map()
let $o-slide-dir := ppt:uri-ppt-slides-dir($s-pres)
let $o-slide-rels-dir := ppt:uri-ppt-slide-rels-dir($s-pres)
let $orig-slide-name := ppt:uri-ppt-slide($s-pres,$s-idx)
let $orig-slide-rels:= ppt:uri-ppt-slide-rels($s-pres,$s-idx)
let $new-slide-name := ppt:uri-ppt-slide((),$start-idx)
let $new-slide-rels := ppt:uri-ppt-slide-rels((),$start-idx)

(: add slide to map :)
let $slide := map:put($smap, $new-slide-name, $orig-slide-name)

let $rels := fn:doc($orig-slide-rels)
let $targets := $rels/rel:Relationships/rel:Relationship/@Target
let $u-targs := for $t in $targets
                let $o-uris := if(fn:matches($t,"slideLayout")) 
                               then $t (: fn:concat(ppt:uri-ppt-dir($s-pres),fn:concat(fn:substring-after($t,"../"))) :)
                               else $t (: fn:replace($t,"\.\./",ppt:uri-ppt-dir($s-pres)) :)
                return $o-uris


let $img-targs := for $u in $u-targs
                  let $image := if(fn:matches($u,"image")) then $u else ()
                  return $image

let $img-count := fn:count($img-targs)
let $new-img-idx := ppt:max-image-id(ppt:uri-ppt-media-dir($t-pres))

(: let $upd-rels := ppt:upd-slide-rels($rels/node(),$img-targs,$new-img-idx) :)
let $upd-rels := ppt:upd-slide-rels($rels,$img-targs,$new-img-idx)


(:add images to map:)
let $images := for $i at $d in $img-targs 
               let $n-idx := ppt:image-id($i)+$new-img-idx
               let $o-img := $i
               let $sfx := fn:substring-after(fn:substring-after($i,"../"),".")
               let $n-img := fn:concat(ppt:uri-ppt-media-dir(()),"image",$n-idx,".",$sfx)
               let $map-update := map:put($smap,$n-img, fn:replace($o-img,"\.\./media/", ppt:uri-ppt-media-dir($s-pres))) 
               return ppt:image-id($i) 


(:ok, put new slide.xml.rels xml in map, then test with instance of before fn:doc and zip for final .pptx :)
let $map-test:= map:put($smap,$new-slide-rels,$upd-rels)
return $smap (:, $images, $upd-rels) :)
 (: $smap :) (: <foo>{$o-slide-dir, $o-slide-rels-dir, $orig-slide-name, $orig-slide-rels, $new-slide-name, $new-slide-rels}</foo> :)
};
(: ====================:)
declare function ppt:check-remove-hm($rels as node())
{
  if(fn:matches($rels/@Target, "handoutMaster")) then () else $rels
};
declare function ppt:passthru-pres-rels($x as node()) as node()*
{
   for $i in $x/node() return ppt:dispatch-pres-rels($i)
};

declare function ppt:dispatch-pres-rels($rels as node()) as node()*
{
      typeswitch($rels)
       case text() return $rels
       case document-node() return document{ppt:passthru-pres-rels($rels)}
       case element(rel:Relationship) return ppt:check-remove-hm($rels) 
       case element(rel:Relationships) return element{fn:name($rels)} {$rels/namespace::*, $rels/@*,passthru-pres-rels($rels)} 
       case element() return  element{fn:name($rels)} {$rels/@*,passthru-pres-rels($rels)}
       default return $rels

};

declare function ppt:remove-hm-from-pres-rels($pres-rels as node())
{
  ppt:dispatch-pres-rels($pres-rels)
};
(: ====================:)

declare function ppt:rel-ids($rels as element(rel:Relationships))
{
   $rels/rel:Relationship/@Id
};
(: ====================:)
(:given a relationships node, and a type (matches on @Target : handout, slide, etc) returns id as integer :)
declare function ppt:rels-rel-id($rels as node(), $type as xs:string*)
{
   (: $rels/r:Relationships/r:Relationship/@Target :)
    let $hmId :=fn:substring-after($rels/rel:Relationships/rel:Relationship[fn:matches(@Target,$type)]/@Id,"rId")
    return if((fn:empty($hmId)) or ($hmId eq "")) then () else xs:integer($hmId)
};
(: ====================:)
declare function ppt:r-id-as-int($rId as xs:string)
{
  xs:integer(fn:substring-after($rId,"rId"))
};
(: ====================:)
(:rename? adjust regardless of hm presence, will adjust slides with insert of new slide :)
declare function ppt:check-adjust-hm($rels as node(), $idx as xs:integer*)
{
if(fn:empty($idx)) then
    let $new-rel := if(fn:not(fn:matches($rels/@Target, "slide")))
                    then
                     let $rId := ppt:r-id-as-int($rels/@Id)
                     return  element{fn:name($rels)} { attribute Id {fn:concat("rId",$rId+1  ) }, $rels/@* except $rels/@Id}
                    else
                     element{fn:name($rels)} { $rels/@* }
    return $new-rel 


else
let $rId := ppt:r-id-as-int($rels/@Id)
(:have to confirm notes master always after slides :)
let $new-rel := if($rId >= $idx or fn:matches($rels/@Target,"notesMasters")) then 
                                       element{fn:name($rels)} { attribute Id {fn:concat("rId",$rId+1  ) }, $rels/@* except $rels/@Id}
                                    else
                                       element{fn:name($rels)} { $rels/@* } 
return $new-rel

};

declare function ppt:passthru-pres-rels-adjust-hm($x as node(), $idx as xs:integer*) as node()*
{
   for $i in $x/node() return ppt:dispatch-pres-rels-adjust-hm($i, $idx)
};

declare function ppt:dispatch-pres-rels-adjust-hm($rels as node(), $hm-id as xs:integer*) as node()*
{

      typeswitch($rels)
       case text() return $rels
       case document-node() return document{ppt:passthru-pres-rels-adjust-hm($rels,$hm-id)}
       case element(rel:Relationship) return ppt:check-adjust-hm($rels, $hm-id) 
       case element(rel:Relationships) return element{fn:name($rels)} {$rels/@*,passthru-pres-rels-adjust-hm($rels,$hm-id)}
       case element() return  element{fn:name($rels)} {$rels/@*,passthru-pres-rels-adjust-hm($rels,$hm-id)}
       default return $rels
};


declare function ppt:pres-rels-adjust-hm($pres-rels as node(), $hm-id as xs:integer*)
{
  ppt:dispatch-pres-rels-adjust-hm($pres-rels, $hm-id)

};


(: ====================:)
declare function ppt:insert-ppt-rels-slide-rel-ORIGINAL($pres-rels as node()*, $start-idx as xs:integer, $hm-id as xs:integer)
{

(: $pres-rels/r:Relationships :)

    let $non-slide-rels := $pres-rels/rel:Relationships/rel:Relationship[fn:not(fn:matches(@Target,"slide\d+\.xml"))]
   
    (:adjust slides: if slide#.xml >= to $start-idx and < $hm-id, increment slide# and rId for slide# :)
    let $orig-slide-rels :=  $pres-rels/rel:Relationships/rel:Relationship[fn:matches(@Target,"slide\d+\.xml")]
    (:adjust here:)

(: let $upd-slide-rels := for $o in $orig-slide-rels :)
                         

 
    
    let $new-slide-rel := element{fn:name($pres-rels/rel:Relationships/rel:Relationship[1])} (:pos don't matter, just name :)  
                                  {attribute Id {fn:concat("rId",1+$start-idx  ) },
                                   attribute Type {"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" },
                                   attribute Target {fn:concat("slides/slide",$start-idx,".xml"  ) }} 


    return element{fn:name($pres-rels/rel:Relationships)} {(($non-slide-rels, $orig-slide-rels, $new-slide-rel))}
};

declare function ppt:insert-ppt-rels-slide-rel($pres-rels as node()*, $start-idx as xs:integer, $hm-id as xs:integer*)
{

(: $pres-rels/r:Relationships :)

    let $non-slide-rels := $pres-rels/rel:Relationships/rel:Relationship[fn:not(fn:matches(@Target,"slide\d+\.xml"))]
   
    (:adjust slides: if slide#.xml >= to $start-idx and < $hm-id, increment slide# and rId for slide# :)
    let $orig-slide-rels :=  $pres-rels/rel:Relationships/rel:Relationship[fn:matches(@Target,"slide\d+\.xml")]
    (:adjust here:)

    let $upd-slide-rels := 
                        for $o in $orig-slide-rels
                        let $slideIdx := xs:integer(fn:substring-before(fn:substring-after($o/@Target, "slides/slide"),".xml"))
                        let $rId := ppt:r-id-as-int($o/@Id)
                         (:may be a problem here with check against handoutmaster id, need to check rId ? instead of slideIdx? :)
                         (:assuming all slides before handout master, this increments rId and slide# for all slides after slide inserted at $start-idx :)
                        let $updSlide := if($slideIdx >= $start-idx) (:and $hm-id > $rId):)  then
                                            let $elem :=  
                                              element{fn:name($pres-rels/rel:Relationships/rel:Relationship[1])} (:pos don't matter, just name :)  
                                              {attribute Id {fn:concat("rId",$rId +1  ) },
                                               attribute Type {"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" },
                                               attribute Target {fn:concat("slides/slide",($slideIdx +1),".xml"  ) }
                                              }
                                            return $elem 
                                         else $o 
                        return $updSlide
                         

 
    
    let $new-slide-rel := element{fn:name($pres-rels/rel:Relationships/rel:Relationship[1])} (:pos don't matter, just name :)  
                                  {attribute Id {fn:concat("rId",1 +$start-idx  ) },
                                   attribute Type {"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" },
                                   attribute Target {fn:concat("slides/slide",$start-idx,".xml"  ) }} 


    return  element{fn:name($pres-rels/rel:Relationships)} {($non-slide-rels, $upd-slide-rels, $new-slide-rel)}
}; 
(: ====================:)
declare function ppt:c-types-remove-theme($ctypes as node(), $theme-ids as xs:string*)
{
   ppt:ct-utils-remove-theme($ctypes, $theme-ids) 
  (:  ppt:dispatch-ct-remove-theme($ctypes) :)
};
(: ====================:)
declare function ppt:c-types-remove-hm($ctypes as node())
{
  ppt:ct-utils-remove-hm($ctypes)
};
(: ================================== :)
declare function ppt:c-types-add-slide($c-types , $slide-idx)
{
   ppt:ct-utils-add-slide($c-types, $slide-idx)
};


(: ================================== :)
declare function ppt:c-types-add-types($c-types as node(), $types as xs:string*)
{
   ppt:ct-utils-add-defaults($c-types, $types)
  (:<fubar>{$types}</fubar> :)
};
(: ================================== :)
(: BEGIN UPDATE FINAL PRESENTATION.XML ================================== :)

declare function ppt:passthru-remove-handoutlst($x as node()) as node()*
{
   for $i in $x/node() return ppt:dispatch-remove-handoutlst($i)
};


declare function ppt:dispatch-remove-handoutlst($pres-xml as node())
{
  typeswitch($pres-xml)
       case text() return $pres-xml
       case document-node() return   document{ppt:passthru-remove-handoutlst($pres-xml)} 
       case element(p:handoutMasterIdLst) return ()
       case element() return  element{fn:name($pres-xml)} {$pres-xml/@*,  $pres-xml/namespace::*,passthru-remove-handoutlst($pres-xml)}
       default return $pres-xml

};


declare function ppt:update-nm($pres-xml as node(), $new-nm-id as xs:string*)
{
   element{fn:name($pres-xml)} {attribute r:id{ $new-nm-id }}

};

declare function ppt:add-sld-ORIG($pres-xml as node(), $new-sld-id as node())
{
  (:$pres-xml:)
 (: need to account for 1- case when two slides have the same id 2-multiple slides will need children rIds updated :)
  let $children := ($pres-xml/node(), $new-sld-id)
  let $ordered := for $c in $children
                  order by $c/@r:id
                  return $c
  return  element{fn:name($pres-xml)}  {$pres-xml/@*, $ordered }
};

declare function ppt:add-sld($pres-xml as node(), $new-sld-id as node())
{
  (:$pres-xml:)
 (: need to account for 1- case when two slides have the same id 2-multiple slides will need children rIds updated :)
  let $children := ($pres-xml/node())
  let $new-sld-rId := ppt:r-id-as-int($new-sld-id/@r:id)
  let $upd-sld-id := 1256 
  let $upd-children := for $c at $n in $children
                       let $rId := ppt:r-id-as-int($c/@r:id)
                       let $slide := if($rId >= $new-sld-rId ) then
                                       let $new-rId := fn:concat("rId",($rId+1))
                                       return  element p:sldId{attribute id {$upd-sld-id + $n (:$c/@id:) } , attribute r:id { $new-rId  } }
                                     else
                                        element p:sldId{attribute id {$upd-sld-id + $n (:$c/@id:) } , attribute r:id { $c/@r:id  } }
                                       (: $c :)

                       return $slide
  let $all-children := ($upd-children, $new-sld-id)              
  let $ordered-c := for $c in $all-children
                    order by $c/@r:id
                    return $c
  return  element{fn:name($pres-xml)}  {$pres-xml/@*, $ordered-c }
};

declare function ppt:passthru-add-slide-id($x as node(), $new-sld-id as node(), $new-nm-id as xs:string*) as node()*
{
   for $i in $x/node() return ppt:dispatch-add-slide-id($i, $new-sld-id, $new-nm-id)
};

declare function ppt:dispatch-add-slide-id($pres-xml as node(), $new-sld-id as node(), $new-nm-id as xs:string*) 
{
  typeswitch($pres-xml)
       case text() return $pres-xml
       case document-node() return   document{ppt:passthru-add-slide-id($pres-xml,$new-sld-id, $new-nm-id)} 
       case element(p:sldIdLst) return ppt:add-sld($pres-xml, $new-sld-id)
       case element(p:notesMasterId) return ppt:update-nm($pres-xml, $new-nm-id)
       case element() return  element{fn:name($pres-xml)} {$pres-xml/@*,  $pres-xml/namespace::*,passthru-add-slide-id($pres-xml,$new-sld-id, $new-nm-id)}
       default return $pres-xml

};

(: declare function ppt:update-pres-xml($pres-xml as node(),$final-pres-rels as node(), $id as xs:integer) :)
(:ppt:update-pres-xml($pres-xml,$final-pres-rels, $s-pres, $start-idx) :)
declare function ppt:update-pres-xml($pres-xml as node(),$final-pres-rels as node(),$src-dir as xs:string, $id as xs:integer)
{
  let $pres-no-hm-lst :=  ppt:dispatch-remove-handoutlst($pres-xml) (: , $final-pres-rels, $id) :)
  let $newid := "256"
  let $slide-xml :=fn:concat("slide",$id,".xml")

  (: original rId of slide in original presentation.xml.rels --to check in presentation.xml-- for slide#.xml :)
  let $src-pres-rel-id := fn:doc(ppt:uri-ppt-rels($src-dir))/rel:Relationships/rel:Relationship[fn:ends-with(@Target,$slide-xml)]/@Id
  (: original id of slide in original presentation.xml for slide#.xml :)
  let $src-pres-slide-id := fn:doc(ppt:uri-ppt-presentation($src-dir))/p:presentation/p:sldIdLst/p:sldId[fn:matches(@r:id,$src-pres-rel-id)]/@id

  (:now check rId to use in $final-pres-rels :)
  let $new-pres-rel-id := $final-pres-rels/rel:Relationship[fn:ends-with(@Target,$slide-xml)]/@Id 

(:could be more than one of these, have to account for :)
  let $new-nm-id := $final-pres-rels/rel:Relationship[fn:ends-with(@Type,"notesMaster")]/@Id  

  (:construct new p:sldId:)
  let $new-sld-id := element p:sldId{attribute id {$src-pres-slide-id } , attribute r:id { $new-pres-rel-id  } }

  

  let $new-pres-xml := ppt:dispatch-add-slide-id($pres-no-hm-lst, $new-sld-id, $new-nm-id)  
  
  return $new-pres-xml (:,$new-sld-id, $new-pres-rel-id, $final-pres-rels, $src-pres-rel-id, $src-pres-slide-id, $slide-xml,$src-dir, $id, $pres-no-hm-lst)
:)
};

declare function ppt:update-pres-xml-ORIG($pres-xml as node(),$final-pres-rels as node(),$src-dir as xs:string, $id as xs:integer)
{
  let $pres-no-hm-lst :=  ppt:dispatch-remove-handoutlst($pres-xml) (: , $final-pres-rels, $id) :)
  let $newid := "256"
  let $slide-xml :=fn:concat("slide",$id,".xml")

  (: original rId of slide in original presentation.xml.rels --to check in presentation.xml-- for slide#.xml :)
  let $src-pres-rel-id := fn:doc(ppt:uri-ppt-rels($src-dir))/rel:Relationships/rel:Relationship[fn:ends-with(@Target,$slide-xml)]/@Id
  (: original id of slide in original presentation.xml for slide#.xml :)
  let $src-pres-slide-id := fn:doc(ppt:uri-ppt-presentation($src-dir))/p:presentation/p:sldIdLst/p:sldId[fn:matches(@r:id,$src-pres-rel-id)]/@id

  (:now check rId to use in $final-pres-rels :)
  let $new-pres-rel-id := $final-pres-rels/rel:Relationship[fn:ends-with(@Target,$slide-xml)]/@Id 

(:could be more than one of these, have to account for :)
  let $new-nm-id := $final-pres-rels/rel:Relationship[fn:ends-with(@Type,"notesMaster")]/@Id  

  (:construct new p:sldId:)
  let $new-sld-id := element p:sldId{attribute id {$src-pres-slide-id } , attribute r:id { $new-pres-rel-id  } }

  

  let $new-pres-xml := ppt:dispatch-add-slide-id($pres-no-hm-lst, $new-sld-id, $new-nm-id)  
  
  return $new-pres-xml (:,$new-sld-id, $new-pres-rel-id, $final-pres-rels, $src-pres-rel-id, $src-pres-slide-id, $slide-xml,$src-dir, $id, $pres-no-hm-lst)
:)
};

(: END UPDATE FINAL PRESENTATION.XML ================================== :)


(:BEGIN  function to merge slide from one deck to another maintaining destination formatting :)
(: $t-pres :="/one_pptx_parts/"    target presentation:)
(: $s-pres :="/two_pptx_parts/"    source presentation:)
(: $s-idx  := 2                    index of slide in source to copy to target :)
(: $start-idx := 2                 insertion index of target presentation :)

declare function ppt:merge-slide($t-pres as xs:string, $s-pres as xs:string, $s-idx as xs:integer, $start-idx as xs:integer)
{
let $t-uris := ppt:package-uris-from-directory($t-pres)   (:uris for target files :)
let $s-uris := ppt:package-uris-from-directory($s-pres)   (:uris for source files :)  

(:removing themes associated with handoutMasters from uris :)
(: final themes get their own map :)
let $uri-handout-master-rels := ppt:uri-ppt-handout-master-rels-dir($t-pres)

(: following returns theme1.xml, theme2.xml, etc. :)
let $theme-ids := ppt:handout-master-theme-ids($uri-handout-master-rels)

let $theme-map := map:map()
let $theme-uris := for $t in $t-uris
                   let $theme-uri := 
                       if(fn:matches($t,"theme\d+\.xml")) then
                          let $check := if(fn:empty($theme-ids)) then
                                        fn:substring-after($t,$t-pres) 
                                        else
                                        for $id in $theme-ids
                                        let $x := if(fn:matches($t,fn:concat($id,"$"))) then () 
                                                 else
                                                 fn:substring-after($t,$t-pres)  
                                       return $x
                          return $check
                       else ()
                   return if(fn:empty($theme-uri))then () else map:put($theme-map,$theme-uri,$t)

(:map for slides and slide relationships (images,etc.) :)

let $slide-map := map:map()
let $new-slide-map := ppt:slide-and-relationships($t-pres, $s-pres, $s-idx, $start-idx)

let $sld-rels-img-types := ppt:sld-rel-image-types($new-slide-map)

(: rest of the uris :)
let $uri-map := map:map()

(:add bulk of uris to map, adjust for new slide#.xml, slide#.xml.rels :)
(: update slide#.xml, slide#.xml.rels accordingly :)
let $final-uris := for $t in $t-uris
                   let $upd-uri := 
                   if(fn:matches($t,"theme\d+\.xml")) then ()        (:themes already in own map :)
                   else if(fn:matches($t,"handoutMaster")) then ()   (:removing hm for now:)
                   else if(fn:ends-with($t,"[Content_Types].xml")) then ()  (:will put in own map X :)
                   else if(fn:ends-with($t,"presentation.xml")) then ()     (:also in map X :)
                   else if(fn:ends-with($t,"presentation.xml.rels")) then () (: also in map X :)
                   
                   else if(fn:matches($t,"slide\d+\.xml$")) then
                     let $slideoriguri := fn:replace($t,"slide\d+\.xml$","")
                     let $newuri := fn:substring-after($slideoriguri,$t-pres)
                     let $slideoname := fn:substring-after($t,$slideoriguri)
                     let $slideidx := fn:substring-before(fn:substring-after($slideoname,"slide",""),".xml","")
                     let $slideint := xs:integer($slideidx)
                     let $final := if($slideint >= $start-idx)
                                                      then fn:concat($newuri,"slide",$slideint+1,".xml")
                                                    else fn:concat($newuri,$slideoname)
                     return $final
                   else if(fn:matches($t,"slide\d+\.xml.rels$")) then
                     let $slideoriguri := fn:replace($t,"slide\d+\.xml.rels$","")
                     let $newuri := fn:substring-after($slideoriguri,$t-pres)
                     let $slideoname := fn:substring-after($t,$slideoriguri)
                     let $slideidx := fn:substring-before(fn:substring-after($slideoname,"slide",""),".xml.rels","")
                     let $slideint := xs:integer($slideidx)
                     let $final := if($slideint >= $start-idx)
                                                      then fn:concat($newuri,"slide",$slideint+1,".xml.rels")
                                                    else fn:concat($newuri,$slideoname)
                     return $final
                   else fn:substring-after($t,$t-pres)                   
                   let $key := if(fn:empty($upd-uri)) then () else
                               map:put($uri-map,$upd-uri,$t)
                   return $key

let $t-pres-rels:= fn:doc(ppt:uri-ppt-rels($t-pres))

(: max-slide-id? or just use idx used forplacement, increment accordingly? :)
(: let $max-slide-id := ppt:max-slide-id(ppt:uri-ppt-slides-dir($t-pres)) :)

(: lets try: add slide t rels, then iterate thru and increase any rId >= to inserted rId of slide, based on this we'll finally update presenation.xml and content-types - then done :)

let $hm-id :=ppt:rels-rel-id($t-pres-rels,"handout")
let $pres-rels-no-hm := ppt:remove-hm-from-pres-rels($t-pres-rels)
let $adjusted-for-hm :=  ppt:pres-rels-adjust-hm($pres-rels-no-hm, $hm-id)


let $final-pres-rels := ppt:insert-ppt-rels-slide-rel($adjusted-for-hm, $start-idx, $hm-id)
  
(: now have to update presentation.xml and content-types :)
let $pres-xml := fn:doc(ppt:uri-ppt-presentation($t-pres))
let $c-types := fn:doc(ppt:uri-content-types($t-pres))/node()
(:need to pass $start-idx to add slide, update other slides if req'd
       pass theme-ids, remove theme-id
       pass image types - if req'd
       pass slidemasters/slidelayouts -next round
:)
let $c-types-no-theme := ppt:c-types-remove-theme($c-types, $theme-ids)
let $c-types-no-hm := ppt:c-types-remove-hm($c-types-no-theme)

(: have to account for slide incrementing here based on where inserted :)
let $upd-ctypes :=  ppt:c-types-add-slide($c-types-no-hm ,$start-idx ) 
(:
let $image-defaults := for $u in $sld-rels-img-types
                       let $ext := $u
                       let $ct := fn:concat("image/",$u)
                       return <Default Extension={$ext} ContentType={$ct}/>
:)
let $final-ctypes := if(fn:empty($sld-rels-img-types)) then $upd-ctypes else
                     ppt:c-types-add-types($upd-ctypes,$sld-rels-img-types)

(:
let $final-ctypes := if(fn:empty($image-defaults)) then $upd-ctypes else
                     ppt:c-types-add-types($upd-ctypes,$image-defaults)
:)

(: need to pass xml/node with slideorig id from source-pres :)

let $final-pres := ppt:update-pres-xml($pres-xml,$final-pres-rels, $s-pres, $s-idx) 

let $pres-root-map := map:map() 
let $uri-pres := ppt:uri-ppt-presentation(())
let $uri-pres-rels := ppt:uri-ppt-rels(())
let $uri-c-types := ppt:uri-content-types(())

let $mapupd1 := map:put( $pres-root-map,$uri-pres, $final-pres)
let $mapupd2 := map:put( $pres-root-map, $uri-pres-rels, $final-pres-rels)
let $mapupd3 := map:put( $pres-root-map, $uri-c-types, $final-ctypes)

let $finalmaps := ($pres-root-map , $theme-map, $new-slide-map, $uri-map )
let $parts := for $m in $finalmaps
              let $keys := map:keys($m)
              return $keys (: fn:count($keys) :)

let $finaldocs := for $p in $parts
                  let $val := map:get($finalmaps, $p)
                  return if($val instance of xs:string) then fn:doc($val) else $val

let $manifest := <parts xmlns="xdmp:zip"> 
   {
    for $i in $parts
    let $part :=  <part>{$i}</part>
    return $part
   }
                 </parts>

let $pptx := xdmp:zip-create($manifest, $finaldocs)
return $pptx
}; 

declare function ppt:sld-rel-image-types($map as map:map)
{
let $tKeys := map:keys($map)
let $rels := for $t in $tKeys
             let $doc := map:get($map,$t)
             let $ret := if($doc instance of xs:string) then () else $doc
             return $ret
let $imgTypes := fn:substring-after(fn:substring-after($rels/rel:Relationships/rel:Relationship[fn:ends-with(@Type,"image")]/@Target,"image"),".")
return $imgTypes
};


(:END  function to merge slide from one deck to another maintaining destination formatting :)
(: ====================:)
(: ====================:)
(: ====================:)
(: ====================:)
(: ====================:)
