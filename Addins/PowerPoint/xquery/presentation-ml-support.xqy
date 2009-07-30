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

declare namespace w="http://schemas.openxmlformats.org/wordprocessingml/2006/main";
declare namespace v="urn:schemas-microsoft-com:vml";
declare namespace ve="http://schemas.openxmlformats.org/markup-compatibility/2006";
declare namespace o="urn:schemas-microsoft-com:office:office";
(: declare namespace r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"; :)
declare namespace r= "http://schemas.openxmlformats.org/package/2006/relationships";
declare namespace m="http://schemas.openxmlformats.org/officeDocument/2006/math";
declare namespace wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
declare namespace w10="urn:schemas-microsoft-com:office:word";
declare namespace wne="http://schemas.microsoft.com/office/word/2006/wordml";
declare namespace pkg="http://schemas.microsoft.com/office/2006/xmlPackage";
declare namespace pic="http://schemas.openxmlformats.org/drawingml/2006/picture";
declare namespace pr="http://schemas.openxmlformats.org/package/2006/relationships";
declare namespace types="http://schemas.openxmlformats.org/package/2006/content-types";
declare namespace zip="xdmp:zip";

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
                 let $theme := $doc/r:Relationships/r:Relationship/@Target
                 let $theme-uri := fn:substring-after($theme,"../theme/")
                 return $theme-uri
  return $theme-idx
};

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
       case document-node() return document {$rels/@*,ppt:passthru-rels($rels, $new-img-idx)}
       case element(r:Relationship) return ppt:update-rels-rel($rels, $new-img-idx) 
       case element() return  element{fn:name($rels)} {$rels/@*,passthru-rels($rels, $new-img-idx)}
       default return $rels

};

declare function ppt:upd-slide-rels($orig-slide-rels as element(r:Relationships),$img-targs as xs:string*,$new-img-idx as xs:integer)
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
let $targets := $rels/r:Relationships/r:Relationship/@Target
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

let $upd-rels := ppt:upd-slide-rels($rels/node(),$img-targs,$new-img-idx)

(:add images to map:)
let $images := for $i at $d in $img-targs 
               let $n-idx := ppt:image-id($i)+$new-img-idx
               let $o-img := $i
               let $sfx := fn:substring-after(fn:substring-after($i,"../"),".")
               let $n-img := fn:concat(ppt:uri-ppt-media-dir(()),"image",$n-idx,".",$sfx)
               let $map-update := map:put($smap,$n-img,fn:replace($o-img,"\.\./",ppt:uri-ppt-media-dir($s-pres)))
               return ppt:image-id($i) 


(:ok, put new slide.xml.rels xml in map, then test with instance of before fn:doc and zip for final .pptx :)
let $map-test:= map:put($smap,$new-slide-rels,$upd-rels)
return $smap (:, $images, $upd-rels) :)
 (: $smap :) (: <foo>{$o-slide-dir, $o-slide-rels-dir, $orig-slide-name, $orig-slide-rels, $new-slide-name, $new-slide-rels}</foo> :)
};

