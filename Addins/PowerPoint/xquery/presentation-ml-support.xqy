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

(:version 1.0-3:)

declare default element namespace "http://schemas.openxmlformats.org/package/2006/relationships";

(: ================== BEGIN serialize Presentation as XML (OPC) ============= :)
declare function ppt:formatbinary(
   $s as xs:string*
) as xs:string*
{
 (:debug test:)(: xdmp:invoke("formatbinary.xqy",( xs:QName("ppt:image"), $s)) :)

 if(fn:string-length($s) > 0) then
     let $firstpart := fn:concat(fn:substring($s,1,76))
      let $tail := fn:substring-after($s,$firstpart) 
      let $tail := fn:substring($s,77) 
     return ($firstpart,ppt:formatbinary($tail))
                  else
             ()

};

declare function ppt:get-part-content-type(
   $uri as xs:string
) as xs:string?
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
   else if(fn:matches($uri, "commentAuthors.xml"))
   then
      "application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml"
   else if(fn:matches($uri, "slideMaster\d+\.xml"))
   then 
      "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"
   else if(fn:matches($uri, "slideLayout\d+\.xml"))
   then
      "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"
   else if(fn:matches($uri,"theme\d+\.xml"))
   then
       "application/vnd.openxmlformats-officedocument.theme+xml"
   else if(fn:matches($uri,"handoutMaster\d+\.xml"))
   then
        "application/vnd.openxmlformats-officedocument.presentationml.handoutMaster+xml"
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
   else if(fn:ends-with(fn:upper-case($uri),"JPEG")) 
   then
      "image/jpeg"
   else if(fn:ends-with(fn:upper-case($uri),"WMF")) 
   then
      "image/x-wmf"
   else if(fn:ends-with(fn:upper-case($uri),"PNG")) 
   then
      "image/png"
   else if(fn:ends-with(fn:upper-case($uri),"GIF"))
   then
       "image/gif"
   else if(fn:matches($uri,"customXml/itemProps\d+\.xml")) then
      "application/vnd.openxmlformats-officedocument.customXmlProperties+xml"
   else if(fn:matches($uri,"customXml/item\d+\.xml")) then
      "application/xml"
   else
       ()
    
};

declare function ppt:get-part-attributes(
   $uri as xs:string
) as node()*
{
  let $cleanuri := fn:replace($uri,"\\","/")
  let $name := attribute pkg:name{$cleanuri}
  let $contenttype := attribute pkg:contentType{ppt:get-part-content-type($cleanuri)}
  let $padding := if(fn:ends-with($cleanuri,".rels")) then

                     if(fn:starts-with($cleanuri,"/ppt/glossary") or 
                        fn:starts-with($cleanuri,"/ppt/slides/_rels") or
                        fn:starts-with($cleanuri,"/ppt/notesSlides/_rels") or
                        fn:starts-with($cleanuri,"/ppt/slideLayouts/_rels") or
                        fn:starts-with($cleanuri,"/ppt/slideMasters/_rels") or
                        fn:starts-with($cleanuri,"/ppt/handoutMasters/_rels")
                       ) then
                         ()
                    
                     else if(fn:starts-with($cleanuri,"/_rels")) then
                      attribute pkg:padding{ "512" }
                     else    
                      attribute pkg:padding{ "256" }
                  else
                     ()
  let $compression := if(fn:ends-with(fn:upper-case($cleanuri),"JPEG") or 
                         fn:ends-with(fn:upper-case($cleanuri),"PNG") or
                         fn:ends-with(fn:upper-case($cleanuri),"GIF")) then 
                         attribute pkg:compression { "store" } 
                      else ()
  
  return ($name, $contenttype, $padding, $compression)
};

declare function ppt:get-package-part(
   $directory as xs:string, 
   $uri as xs:string
) as node()?
{
  let $fulluri := $uri
  let $docuri := fn:concat("/",fn:substring-after($fulluri,$directory))
  let $data := fn:doc($fulluri)

  let $part := if(fn:empty($data) or fn:ends-with($fulluri,"[Content_Types].xml")) then () 
               else 
                 if(fn:ends-with(fn:upper-case($fulluri),".JPEG") or 
                    fn:ends-with(fn:upper-case($fulluri),".WMF") or 
                    fn:ends-with(fn:upper-case($fulluri),".GIF") or 
                    fn:ends-with(fn:upper-case($fulluri),".PNG")) then
                   let $bin := xs:base64Binary(xs:hexBinary($data)) cast as xs:string 
                    let $formattedbin := fn:string-join(ppt:formatbinary($bin),"&#x9;&#xA;") 
                    return  element pkg:part { ppt:get-part-attributes($docuri), element pkg:binaryData { $formattedbin  }   }
                 else
                   element pkg:part { ppt:get-part-attributes($docuri), element pkg:xmlData { $data }}
  return  $part 
};

declare function ppt:package-make(
   $directory as xs:string, 
   $uris as xs:string*
) as element(pkg:package)
{
  	let $package := element pkg:package { 
                            for $uri in $uris
                            let $part := ppt:get-package-part($directory,$uri)
                            return $part }
                           
	return $package

(: processing instructions generated when Word or PPT 'Save As' XML:)
(: not currently required for Office to open file :)
(: <?mso-application progid="Word.Document"?>, $package :)
(: <?mso-application progid="PowerPoint.Show"?> :)

};
(: ================== END serialize Presentation as XML (OPC) ============= :)

declare function ppt:directory-uris(
   $docuri as xs:string
) as xs:string*
{
  	cts:uris("","document",cts:directory-query($docuri,"infinity"))
};

declare function ppt:directory-uris(
   $docuri as xs:string, 
   $depth as xs:string
) as xs:string*
{
  	cts:uris("","document",cts:directory-query($docuri,$depth))
};

declare function ppt:package-files-only(
   $uris as xs:string*
) as xs:string*
{
  	for $uri in $uris
  	let $u := if(fn:ends-with($uri,"/")) then () else $uri
  	return $u
};

(: ================== BEGIN file and directory URI helper functions  ======== :)

(: /foo_pptx_parts/ppt/slides/slide2.xml => /foo_PNG/Slide2.PNG :)
declare function ppt:uri-slide-xml-to-slide-png(
   $uri as xs:string
) as xs:string
{     
        fn:replace($uri,"^(.*)(_pptx_parts/ppt/slides/)slide(\d+).xml$","$1_PNG/Slide$3.PNG")
};

(: /foo_PNG/Slide2.PNG => /foo_pptx_parts/ppt/slides/slide2.xml :)
declare function ppt:uri-slide-png-to-slide-xml(
   $uri as xs:string
) as xs:string
{
       fn:replace($uri,"^(.*)(_PNG/)Slide(\d+).PNG$","$1_pptx_parts/ppt/slides/slide$3.xml")
};

declare function ppt:uri-content-types(
   $dir as xs:string?
) as xs:string
{
	fn:concat($dir,"[Content_Types].xml")
};

declare function ppt:uri-rels-dir(
   $dir as xs:string?
) as xs:string
{
  	fn:concat($dir,"_rels/")
};

declare function ppt:uri-rels(
   $dir as xs:string?
) as xs:string
{
    	fn:concat(ppt:uri-rels-dir($dir),".rels") 
};

declare function ppt:uri-docprops-dir(
   $dir as xs:string?
) as xs:string
{
    	fn:concat($dir,"docProps/")
};

declare function ppt:uri-app-props(
   $dir as xs:string?
) as xs:string
{
      	fn:concat(ppt:uri-docprops-dir($dir),"app.xml")
};

declare function ppt:uri-core-props(
   $dir as xs:string?
) as xs:string
{
     	 fn:concat(ppt:uri-docprops-dir($dir),"core.xml")
};

declare function ppt:uri-ppt-dir(
   $dir as xs:string?
) as xs:string
{
      	fn:concat($dir, "ppt/")
};

declare function ppt:uri-ppt-rels-dir(
   $dir as xs:string?
) as xs:string
{
      	fn:concat(ppt:uri-ppt-dir($dir),"_rels/")
};

declare function ppt:uri-ppt-rels(
   $dir as xs:string?
) as xs:string
{
     	fn:concat(ppt:uri-ppt-rels-dir($dir),"presentation.xml.rels")
};

declare function ppt:uri-ppt-handout-masters-dir(
   $dir as xs:string?
) as xs:string
{
      	fn:concat(ppt:uri-ppt-dir($dir),"handoutMasters/")
};

declare function ppt:uri-ppt-handout-master-rels-dir(
   $dir as xs:string?
) as xs:string
{
      	fn:concat(ppt:uri-ppt-handout-masters-dir($dir),"_rels/")
};

declare function ppt:uri-ppt-handout-master(
   $dir as xs:string?, 
   $idx as xs:integer
) as xs:string
{
    	let $handoutMasterFile := fn:concat("handoutMaster",$idx,".xml")
   	return fn:concat(ppt:uri-ppt-handout-masters-dir($dir), $handoutMasterFile)
};

declare function ppt:uri-ppt-handout-master-rels(
   $dir as xs:string?, 
   $idx as xs:integer
) as xs:string
{
    	let $handoutMasterRelsFile := fn:concat("handoutMaster",$idx,".xml.rels")
    	return fn:concat( ppt:uri-ppt-handout-master-rels-dir($dir),$handoutMasterRelsFile)
};

declare function ppt:uri-ppt-media-dir(
   $dir as xs:string?
) as xs:string
{
     	fn:concat(ppt:uri-ppt-dir($dir),"media/")
};

declare function ppt:uri-ppt-notes-masters-dir(
   $dir as xs:string?
) as xs:string
{
     	fn:concat(ppt:uri-ppt-dir($dir),"notesMasters/")
};

declare function ppt:uri-ppt-notes-masters-rels-dir(
   $dir as xs:string?
) as xs:string
{
    	fn:concat(ppt:uri-ppt-notes-masters-dir($dir),"_rels/")
}; 

declare function ppt:uri-ppt-notes-master(
   $dir as xs:string?, 
   $idx as xs:integer
) as xs:string
{
    	let $notesMasterFile := fn:concat("notesMaster",$idx,".xml")
    	return fn:concat(ppt:uri-ppt-notes-masters-dir($dir), $notesMasterFile)
};

declare function ppt:uri-ppt-notes-master-rels(
   $dir as xs:string?, 
   $idx as xs:integer
) as xs:string
{
    	let $notesMasterRelsFile := fn:concat("notesMaster",$idx,".xml.rels")
    	return fn:concat( ppt:uri-ppt-notes-masters-rels-dir($dir),$notesMasterRelsFile)
};

declare function ppt:uri-ppt-slide-layouts-dir(
   $dir as xs:string?
) as xs:string
{
     	fn:concat(ppt:uri-ppt-dir($dir),"slideLayouts/")
};

declare function ppt:uri-ppt-slide-layout-rels-dir(
   $dir as xs:string?
) as xs:string
{
    	fn:concat(ppt:uri-ppt-slide-layouts-dir($dir),"_rels/")
}; 

declare function ppt:uri-ppt-slide-layout(
   $dir as xs:string?, 
   $idx as xs:integer
) as xs:string
{
    	let $slideLayoutFile := fn:concat("slideLayout",$idx,".xml")
    	return fn:concat(ppt:uri-ppt-slide-layouts-dir($dir), $slideLayoutFile)
};

declare function ppt:uri-ppt-slide-layout-rels(
   $dir as xs:string?, 
   $idx as xs:integer
) as xs:string
{
    	let $slideLayoutRelsFile := fn:concat("slideLayout",$idx,".xml.rels")
    	return fn:concat( ppt:uri-ppt-slide-layout-rels-dir($dir),$slideLayoutRelsFile)
};

declare function ppt:uri-ppt-slide-masters-dir(
   $dir as xs:string?
) as xs:string
{
     	fn:concat(ppt:uri-ppt-dir($dir),"slideMasters/")
};

declare function ppt:uri-ppt-slide-master-rels-dir(
   $dir as xs:string?
) as xs:string
{
    	fn:concat(ppt:uri-ppt-slide-masters-dir($dir),"_rels/")
}; 

declare function ppt:uri-ppt-slide-master(
   $dir as xs:string?, 
   $idx as xs:integer
) as xs:string
{
    	let $slideMasterFile := fn:concat("slideMaster",$idx,".xml")
   	return fn:concat(ppt:uri-ppt-slide-masters-dir($dir), $slideMasterFile)
};

declare function ppt:uri-ppt-slide-master-rels(
   $dir as xs:string?, 
   $idx as xs:integer
) as xs:string
{
    	let $slideMasterRelsFile := fn:concat("slideMaster",$idx,".xml.rels")
    	return fn:concat( ppt:uri-ppt-slide-master-rels-dir($dir),$slideMasterRelsFile)
};

declare function ppt:uri-ppt-slides-dir(
   $dir as xs:string?
) as xs:string
{
     	fn:concat(ppt:uri-ppt-dir($dir),"slides/")
};


declare function ppt:uri-ppt-slide-rels-dir(
   $dir as xs:string?
) as xs:string
{
    	fn:concat(ppt:uri-ppt-slides-dir($dir),"_rels/")
}; 

declare function ppt:uri-ppt-slide(
   $dir as xs:string?, 
   $idx as xs:integer
) as xs:string
{
    	let $slideFile := fn:concat("slide",$idx,".xml")
    	return fn:concat(ppt:uri-ppt-slides-dir($dir), $slideFile)
};

declare function ppt:uri-ppt-slide-rels(
   $dir as xs:string?, 
   $idx as xs:integer
) as xs:string
{
    	let $slideRelsFile := fn:concat("slide",$idx,".xml.rels")
    	return fn:concat( ppt:uri-ppt-slide-rels-dir($dir),$slideRelsFile)
};

declare function ppt:uri-ppt-theme-dir(
   $dir as xs:string?
) as xs:string
{
     	fn:concat(ppt:uri-ppt-dir($dir),"theme/")
};

declare function ppt:uri-ppt-theme-rels-dir(
   $dir as xs:string?
) as xs:string
{
    	fn:concat(ppt:uri-ppt-theme-dir($dir),"_rels/")
}; 

declare function ppt:uri-ppt-theme(
   $dir as xs:string?, 
   $idx as xs:integer
) as xs:string
{
    	let $themeFile := fn:concat("theme",$idx,".xml")
    	return fn:concat(ppt:uri-ppt-theme-dir($dir), $themeFile)
};

declare function ppt:uri-ppt-theme-rels(
   $dir as xs:string?, 
   $idx as xs:integer
) as xs:string
{
    	let $themeRelsFile := fn:concat("theme",$idx,".xml.rels")
    	return fn:concat( ppt:uri-ppt-theme-rels-dir($dir),$themeRelsFile)
};

declare function ppt:uri-ppt-presentation(
   $dir as xs:string?
) as xs:string
{
     	fn:concat(ppt:uri-ppt-dir($dir),"presentation.xml")
};

declare function ppt:uri-ppt-pres-props(
   $dir as xs:string?
) as xs:string
{
     	fn:concat(ppt:uri-ppt-dir($dir),"presProps.xml")
};

declare function ppt:uri-ppt-table-styles(
   $dir as xs:string?
) as xs:string
{
     	fn:concat(ppt:uri-ppt-dir($dir),"tableStyles.xml")
};

declare function ppt:uri-ppt-view-props(
   $dir as xs:string?
) as xs:string
{
     	fn:concat(ppt:uri-ppt-dir($dir),"viewProps.xml")
};
(: ================== END file and directory URI helper functions  ========== :)

declare function ppt:uri-ppt-handout-master-rels-map(
   $pkg-map as map:map
) as xs:string*
{
   	let $keys := map:keys($pkg-map)
   	let $hm-dir :=  ppt:uri-ppt-handout-master-rels-dir(())
   	return for $k in $keys
           	    let $hm := 	if(fn:matches($k, fn:concat($hm-dir,"handoutMaster\d+\.xml.rels?"))) then 
					map:get($pkg-map,$k) 
				else ()
                    return $hm
};

declare function ppt:map-update(
   $target-map as map:map, 
   $src-map as map:map
) as map:map*
{
         for $k in  map:keys($src-map)
         let $kval := map:get($src-map, $k)
         return map:put($target-map, $k, $kval)
};

declare function ppt:map-max-image-id(
   $pkg as map:map
) as xs:integer
{
   	let $keys := map:keys($pkg)
   	let $numbers := 
		 for $k in $keys
                 return if(fn:matches($k, "image")) then 
                             xs:integer(fn:substring-before(fn:substring-after($k,"image"),"."))
                        else ()
   	return if(fn:empty($numbers)) then 1 else fn:max($numbers)
};

(:
declare function ppt:max-file-id(
   $dir as xs:string*, 
   $type as xs:string*, 
   $depth as xs:string
) as xs:integer
{  (:fn:max on inner flwr, default return is 0 :)
  	let $files :=  ppt:directory-uris($dir,"1")
  	let $numbers := 
                  if(fn:empty($files)) then 0
                  else
                     for $i in $files
                     let $tmp1 := fn:substring-after($i,fn:concat($dir,$type))
                     let $tmp2 := fn:substring-before($tmp1,".")
                     return xs:integer($tmp2)
  	return fn:max($numbers)

};

declare function ppt:max-image-id(
   $dir as xs:string*
) as xs:integer
{
  
  	ppt:max-file-id($dir, "image", "infinity") 

};

declare function ppt:max-slide-id(
   $dir as xs:string*
) as xs:integer
{
  	ppt:max-file-id($dir, "slide", "1") 
};
:)

declare function ppt:handout-master-theme-index(
   $hm-rels as xs:string*
) as xs:string*
{       (: xs:integer cast is not reqd, being used in concat :)
	for $rel in $hm-rels
        let $doc := fn:doc($rel)
        let $theme := $doc/rel:Relationships/rel:Relationship/@Target
        let $theme-uri := fn:substring-before(fn:substring-after($theme,"../theme/theme"),".xml")
        return $theme-uri           
};

(:may want a generic fileid function, similar to max :)
declare function ppt:image-id(
   $uri as xs:string
) as xs:integer
{
  	xs:integer(fn:substring-before(fn:substring-after($uri,"image"),"."))
};

(: ================== BEGIN Update slide#.xml.rels  ========================= :)
declare function ppt:update-rels-relationship(
   $r as node(), 
   $n-idx as xs:integer
) as node()*
{
(:currently only updates for images :)
   	if(fn:matches($r/@Target,"slideLayout")) then $r
        else if(fn:matches($r/@Target,"notesSlide\d+\.xml?")) then () 
   	else if(fn:matches($r/@Target,"image")) then
    		let $target := $r/@Target
    		let $prfx := fn:substring-before($target,"image")
    		let $sfx := fn:substring-after($target,".")
    		let $id := ppt:image-id($target)
    		let $n-targ := fn:replace($target, xs:string($id),xs:string($id + $n-idx))
    		return  element{fn:QName("http://schemas.openxmlformats.org/package/2006/relationships","Relationship")} {$r/@* except $r/@Target, attribute Target{$n-targ}}
   	else $r
};

declare function ppt:passthru-rels(
   $x as node(), 
   $idx as xs:integer
) as node()*
{
   	for $i in $x/node() return ppt:dispatch-slide-rels($i, $idx)
};

declare function ppt:dispatch-slide-rels(
   $rels as node(), 
   $new-img-idx as xs:integer
) as node()*
{
       typeswitch($rels)
        case text() return $rels
        case document-node() return ppt:passthru-rels($rels, $new-img-idx)
        case element(rel:Relationship) return ppt:update-rels-relationship($rels, $new-img-idx) 
        case element(rel:Relationships) return element{fn:QName("http://schemas.openxmlformats.org/package/2006/relationships","Relationships")} {$rels/namespace::*, $rels/@*,passthru-rels($rels, $new-img-idx)}
        case element() return  element{fn:node-name($rels)} {$rels/@*,passthru-rels($rels, $new-img-idx)}
       default return $rels

};

declare function ppt:update-slide-rels(
   $orig-slide-rels as node(),
   $img-targs as xs:string*,
   $new-img-idx as xs:integer
) as element(rel:Relationships)
{
  	ppt:dispatch-slide-rels($orig-slide-rels, $new-img-idx)
};
(: ================== END Update slide#.xml.rels  =========================== :)

declare function ppt:map-slide-and-relationships(
   $to-pkg-map as map:map, 
   $from-pres as xs:string, 
   $from-idx as xs:integer, 
   $to-idx as xs:integer
) as map:map
{
     (:loop through and increment slide#.xml, slide#.xml.rels -keys and values in map 
       reset values of existing keys, for 1+, will add new value, otherwise resets
       existing (with new slide info as value) :)  
        let $to-uris := map:keys($to-pkg-map)
        let $orig-slide-name := ppt:uri-ppt-slide($from-pres,$from-idx)
        let $orig-slide-rels := ppt:uri-ppt-slide-rels($from-pres,$from-idx)
        let $new-slide-name :=  ppt:uri-ppt-slide((),$to-idx)
        let $new-slide-rels :=  ppt:uri-ppt-slide-rels((),$to-idx)

        let $rels := fn:doc($orig-slide-rels)
        let $targets := $rels/rel:Relationships/rel:Relationship/@Target

  	let $img-targs := 
                  for $u in $targets
                  return if(fn:matches($u,"image")) then $u else ()

        let $new-img-idx := ppt:map-max-image-id($to-pkg-map)

     (: update slide#.xml.rels :)
        let $upd-rels := ppt:update-slide-rels($rels,$img-targs,$new-img-idx)

     (:add slide associated images to map :)
        let $images := 
               for $i at $d in $img-targs 
               let $n-idx := ppt:image-id($i)+$new-img-idx
               let $sfx := fn:substring-after(fn:substring-after($i,"../"),".")
               let $n-img := fn:concat(ppt:uri-ppt-media-dir(()),"image",$n-idx,".",$sfx)
               let $map-update := map:put($to-pkg-map,$n-img, fn:replace($i,"\.\./media/", ppt:uri-ppt-media-dir($from-pres))) 
               return ppt:image-id($i)

        let $slide-uris := for $x at $d in $to-uris
                           let $doc := if(fn:matches($x,"/slide\d+\.xml$")) then
                                                $x
                                              else
                                               () 
                           return $doc

        let $slide-rel-uris := for $y in $to-uris
                               let $s-rel-uri := if(fn:matches($y,"slide\d+\.xml.rels$")) then
                                                $y
                                               else
                                                ()
                               return $s-rel-uri
      
     (:loop thru uris, generate updated key, val pairs for slide#.xml docs
       and add to tmp map to be used for updating package map :)
        let $upd-map1 := map:map()
        let $slide-upd-map := 
                            for $s in $slide-uris
                            (: increment keys, not values as values point to the actual urls we want to use 
                               for the updated references :) 
                            let $key-idx := xs:integer(fn:substring-before(fn:substring-after($s,"slides/slide"),".xml"))
                            let $orig-slide-val := map:get($to-pkg-map,$s) 
                            let $final-key     := if($key-idx >= $to-idx) then
                                                     fn:concat(fn:replace($s,"slide\d+\.xml$",""),"slide",($key-idx+1),".xml")
                                                 else $s
                            return map:put($upd-map1 , $final-key, $orig-slide-val)

     (:update package map with updated slide#.xml references :)
        let $slide-upd := ppt:map-update($to-pkg-map, $upd-map1)
   
     (:loop thru uris, generate updated key, val pairs for slide#.xml.rels docs
       and add to tmp map to be used for updating package map :)
        let $upd-map2 := map:map()
        let $slide-rel-upd-map := 
                              for $sr in $slide-rel-uris
                              let $key-idx := xs:integer(fn:substring-before(fn:substring-after($sr,"_rels/slide"),".xml.rels"))
                              let $orig-slide-rel-val := map:get($to-pkg-map,$sr)
                              let $final-rel-key     := if($key-idx >= $to-idx) then
                                                           fn:concat(fn:replace($sr,"slide\d+\.xml.rels$",""),"slide",($key-idx+1),".xml.rels")
                                                        else $sr
                              return map:put($upd-map2 , $final-rel-key, $orig-slide-rel-val)

     (:update package map with updated slide#.xml.rels references :)
        let $slide-rel-upd := ppt:map-update($to-pkg-map, $upd-map2)
   
     (:add updated slide#.xml/slide#.xml.rels references, and slide to be inserted to package map:)
        let $map-rels :=  map:put($to-pkg-map,$new-slide-rels,$upd-rels)
        let $map-slide := map:put($to-pkg-map, $new-slide-name, $orig-slide-name)

	return $to-pkg-map
};

(:removes handout master from presentation.xml.rels :)
declare function ppt:remove-hm-from-pres-rels(
   $pres-rels as element(rel:Relationships) 
) as element(rel:Relationships)
{
     	let $upd-children := for $c in $pres-rels/Relationship
        	             let $rel := if(fn:matches($c/@Target, "handoutMaster")) then () else $c
                	     return $rel
     	return element{fn:QName("http://schemas.openxmlformats.org/package/2006/relationships","Relationships")} 
			   {$pres-rels/@*, $upd-children} 

};

declare function ppt:rel-ids(
   $rels as element(rel:Relationships)
) as xs:string*
{
   	$rels/rel:Relationship/@Id
};

(:given a relationships node, and a type ,matches on @Target : handout, slide, etc,
  function returns id as integer :)
declare function ppt:rels-rel-id(
   $rels as node(), 
   $type as xs:string*
) as xs:integer
{
    	let $hmId :=fn:substring-after($rels/rel:Relationships/rel:Relationship[fn:matches(@Target,$type)]/@Id,"rId")
    	return if((fn:empty($hmId)) or ($hmId eq "")) then () else xs:integer($hmId)
};

declare function ppt:r-id-as-int(
   $rId as xs:string
) as xs:integer
{
  	xs:integer(fn:substring-after($rId,"rId"))
};

declare function ppt:ppt-rels-insert-slide(
   $pres-rels as node(), 
   $start-idx as xs:integer
) as element(rel:Relationships)
{
(:pos don't matter, just name :)   
(:incrementing by one here assumes one slideMaster, should query count of masters, then increment accordingly for new-r-id:)
        let $new-r-id := 1 + $start-idx  (:change to + count of slidemaster :)
        (:if rId >= $new-r-id then increment :)
    	let $non-slide-rels := $pres-rels/Relationship[fn:not(fn:matches(@Target,"slide\d+\.xml"))]
   
        (:adjust slides: if slide#.xml >= to $start-idx, increment slide# and rId for slide# :)
    	let $orig-slide-rels :=  $pres-rels/Relationship[fn:matches(@Target,"slide\d+\.xml")]
    	let $new-slide-rel := element{fn:QName("http://schemas.openxmlformats.org/package/2006/relationships","Relationship")} 
                                  {attribute Id {fn:concat("rId",$new-r-id  ) },
                                   attribute Type {"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" },
                                   attribute Target {fn:concat("slides/slide",$start-idx,".xml"  ) }}

        let $new-non-slide-rels := for $n in $non-slide-rels
                                   let $rId := ppt:r-id-as-int($n/@Id)
                                   return if($rId >= $new-r-id) then 
		  	                        element{fn:QName("http://schemas.openxmlformats.org/package/2006/relationships","Relationship")} 
                                                       { attribute Id {fn:concat("rId",$rId+1  ) }, $n/@* except $n/@Id }
                                          else $n 

        let $new-slide-rels := for $o in $orig-slide-rels
                               let $slideIdx := xs:integer(fn:substring-before(fn:substring-after($o/@Target, "slides/slide"),".xml"))
                               let $rId := ppt:r-id-as-int($o/@Id)
                               return if($slideIdx >= $start-idx)  then
                                                 element{fn:QName("http://schemas.openxmlformats.org/package/2006/relationships","Relationship")} 
                                                 {attribute Id {fn:concat("rId",$rId +1  ) },
                                                  attribute Type {"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" },
                                                  attribute Target {fn:concat("slides/slide",($slideIdx +1),".xml"  ) }
                                                 }
                                      else $o 

        return element{fn:QName("http://schemas.openxmlformats.org/package/2006/relationships","Relationships")} 
                      {($new-non-slide-rels, $new-slide-rels, $new-slide-rel)} 
};

declare function ppt:update-content-types(
   $content-types as node(), 
   $slide-idx as xs:integer,
   $img-types as xs:string*, 
   $theme-ids as xs:string*
) as element(types:Types)
{
        ppt:ct-utils-update-types($content-types, $slide-idx, $img-types, $theme-ids)
};

(: ================== BEGIN update presentation.xml ========================= :)
declare function ppt:passthru-remove-handoutlst(
   $x as node()*
) as node()*
{
   	for $i in $x/node() return ppt:dispatch-remove-handoutlst($i)
};


declare function ppt:dispatch-remove-handoutlst(
   $pres-xml as node()*
) as node()* 
{
       typeswitch($pres-xml)
	case text() return $pres-xml
       	case document-node() return   document{ppt:passthru-remove-handoutlst($pres-xml)} 
       	case element(p:handoutMasterIdLst) return ()
       	case element() return  element{fn:name($pres-xml)} {$pres-xml/@*,  $pres-xml/namespace::*,passthru-remove-handoutlst($pres-xml)}
       default return $pres-xml

};

declare function ppt:update-notesmaster-id(
   $pres-xml as node(), 
   $new-nm-id as xs:string*
) as element(p:notesMasterId)
{
  	element{fn:QName("http://schemas.openxmlformats.org/presentationml/2006/main","p:notesMasterId")} {attribute r:id{ $new-nm-id }}
};

declare function ppt:add-slide(
   $pres-xml as node(), 
   $new-sld-id as node()
) as element(p:sldIdLst)
{
  	let $children := $pres-xml/node()
 	let $new-sld-rId := ppt:r-id-as-int($new-sld-id/@r:id)

        (:take slide id for slide being added, make all one greater as we add back to presentation.xml :)
  	let $upd-sld-id := xs:integer($new-sld-id/@id) + 1
  	let $upd-children := 
                       for $c at $n in $children
                       let $rId := ppt:r-id-as-int($c/@r:id)
                       return if($rId >= $new-sld-rId ) then
                                  let $new-rId := fn:concat("rId",($rId+1))
                                  return  element p:sldId{attribute id {$upd-sld-id + $n } , attribute r:id { $new-rId  } }
                              else
                                  element p:sldId{attribute id {$upd-sld-id + $n } , attribute r:id { $c/@r:id  } }

  	let $all-children := ($upd-children, $new-sld-id)              
  	let $ordered-c := for $c in $all-children 
        	          order by xs:integer(fn:substring($c/@r:id,4))
                          return $c
  	return element{fn:QName("http://schemas.openxmlformats.org/presentationml/2006/main","p:sldIdLst")}  {$pres-xml/@*, $ordered-c } 

};

declare function ppt:passthru-add-slide-id(
   $x as node(), 
   $new-sld-id as node(), 
   $new-nm-id as xs:string*
) as node()*
{
   	for $i in $x/node() return ppt:dispatch-add-slide-id($i, $new-sld-id, $new-nm-id)
};

declare function ppt:dispatch-add-slide-id(
   $pres-xml as node(), 
   $new-sld-id as node(), 
   $new-nm-id as xs:string*
) as node()* 
{
       typeswitch($pres-xml)
     	case text() return $pres-xml
       	case document-node() return  ppt:passthru-add-slide-id($pres-xml,$new-sld-id, $new-nm-id)
       	case element(p:sldIdLst) return ppt:add-slide($pres-xml, $new-sld-id)
       	case element(p:notesMasterId) return ppt:update-notesmaster-id($pres-xml, $new-nm-id)
       	case element() return  element{fn:name($pres-xml)} {$pres-xml/@*,  $pres-xml/namespace::*,passthru-add-slide-id($pres-xml,$new-sld-id, $new-nm-id)}
       default return $pres-xml

};

declare function ppt:update-presentation-xml(
   $presentation-xml as node(),
   $final-pres-rels as node(),
   $src-dir as xs:string, 
   $from-idx as xs:integer,
   $to-idx as xs:integer
) as element(p:presentation)
{
  	let $pres-no-hm-lst :=  ppt:dispatch-remove-handoutlst($presentation-xml)

  	let $slide-xml :=fn:concat("slide",$from-idx,".xml")
        let $new-slide-xml := fn:concat("slide",$to-idx,".xml")

        (: original rId of slide in original presentation.xml.rels --to check in presentation.xml-- for slide#.xml :)
  	let $src-pres-rel-id := fn:doc(ppt:uri-ppt-rels($src-dir))/rel:Relationships/rel:Relationship[fn:ends-with(@Target,$slide-xml)]/@Id

        (: original id of slide in original presentation.xml for slide#.xml :)
  	let $src-pres-slide-id := fn:doc(ppt:uri-ppt-presentation($src-dir))/p:presentation/p:sldIdLst/p:sldId[@r:id eq $src-pres-rel-id]/@id

        (:now check rId to use in $final-pres-rels :)
  	let $new-pres-rel-id := $final-pres-rels/rel:Relationship[fn:ends-with(@Target,$new-slide-xml)]/@Id 

        (:could be more than one of these, have to account for :)
  	let $new-nm-id := $final-pres-rels/rel:Relationship[fn:ends-with(@Type,"notesMaster")]/@Id  

        (:construct new p:sldId:)
        let $new-sld-id := element p:sldId{attribute id {$src-pres-slide-id } , attribute r:id { $new-pres-rel-id  } } 

	let $new-pres-xml := ppt:dispatch-add-slide-id($pres-no-hm-lst, $new-sld-id, $new-nm-id)  
  
  	return $new-pres-xml 

};

(: ================== END update presentation.xml =========================== :)

declare function ppt:slide-index-error()
{
   	fn:error("SlideIndexOutofBounds: ","The index specified for the presentation does not exist.")
};

declare function ppt:list-length-error()
{
        fn:error("ListLengthsNotEqual: ","The lengths of the lists that are dependant on each other differ.") 
};

declare function ppt:validate-list-length-equal(
   $list1 as xs:string+, 
   $list2 as xs:integer+
) as xs:boolean
{
  	fn:count($list1) eq fn:count($list2)
};

declare function ppt:validate-slide-indexes-map(
   $t-map as map:map, 
   $from-pres as xs:string, 
   $from-idx as xs:integer, 
   $insert-index as xs:integer
) as xs:boolean
{

   (:may want to break slide count from map out into own function :)
   	let $keys := map:keys($t-map)
   	let $slides-dir :=  ppt:uri-ppt-slides-dir(())
   	let $slides := for $k in $keys
        	       let $s := if(fn:matches($k, fn:concat($slides-dir,"slide\d+\.xml?"))) then $k else ()
                       return $s
   	let $tconfirm-cnt := fn:count($slides)
 
   	let $s-slide-confirm := ppt:uri-ppt-slides-dir($from-pres)
   	let $sconfirm-cnt := fn:count(ppt:directory-uris($s-slide-confirm,"1"))
   	let $test :=    if($from-idx = 0 or $insert-index = 0) 
        	        then 
                	   	fn:false()
               	        else if($tconfirm-cnt = 0 or $sconfirm-cnt = 0) 
               		then
                    		fn:false() 
              		else if((($insert-index) > $tconfirm-cnt+1 ) or ($from-idx > $sconfirm-cnt ))  
               		then 
                   		fn:false() 
               		else 
                   		fn:true()
   	return $test
};

declare function ppt:slide-rel-image-types(
   $map as map:map
) as xs:string*
{
	let $tKeys := map:keys($map)
	let $rels := for $t in $tKeys
        	     let $doc := map:get($map,$t)
                     return if($doc instance of xs:string) then () else $doc

	let $imgTypes := for $r in $rels
                         let $targs := $r/rel:Relationships/rel:Relationship[fn:ends-with(@Type,"image")]/@Target
                         let $type := for $t in $targs
                                      return fn:substring-after(fn:substring-after($t,"image"),".")
                         return fn:distinct-values($type)

	return if($imgTypes eq "") then 
                  () 
               else $imgTypes
};

 
(:BEGIN  function to merge slide from one deck to another maintaining destination formatting :)
(: $to-pkg-map :=      target presentation, use ppt:package-map to create intial, or pass in existing :)
(: $from-pres :="/two_pptx_parts/"    presentation slide will be merged from            :)
(: $from-idx  := 2                    index of slide in from preso to copy to target    :)
(: $insert-idx := 2                   insertion index of slide merged into preso in map :)

declare function ppt:merge-slide-util(
   $to-pkg-map as map:map?,
   $from-pres as xs:string, 
   $from-idx as xs:integer, 
   $insert-idx as xs:integer
) as map:map 
{

        let $to-uris := map:keys($to-pkg-map)                   (:uris for target presentation files    :)
        let $from-uris := ppt:directory-uris($from-pres)        (:uris for presentation files with slides to insert  :) 
        let $uri-handout-master-rels := ppt:uri-ppt-handout-master-rels-map($to-pkg-map) (:return handoutMaster uri from map :)
        let $theme-ids := ppt:handout-master-theme-index($uri-handout-master-rels) (:return indices for themes related to handoutMaster - 1,2..N :)

        (:remove themes associated with handoutmaster from map:)
 	let $theme-uris := 
                   for $t in $to-uris
                   return if(fn:matches($t,"theme\d+\.xml$")) then
                             for $id in $theme-ids
                             return if(fn:matches($t,fn:concat("theme",$id,".xml"))) then 
                                          map:delete($to-pkg-map,$t)
                                     else 
                                           ()
                          else ()
         
        (:remove handoutMaster :)      
        let $remove-hms := for $to in $to-uris
                           return if(fn:matches($to,"handoutMaster")) then 
                                     map:delete($to-pkg-map, $to)
                           else () 
        
        (:inserts new slide, slide.xml.rels, and any associated images, 
          adjusts references for other slide.xml, slide.xml.rels, and images :)
        let $new-slide-map := ppt:map-slide-and-relationships($to-pkg-map, $from-pres, $from-idx, $insert-idx)
        
        let $sld-rels-img-types := ppt:slide-rel-image-types($new-slide-map)

        (:get presentation.xml.rels for update :)
        let $to-pres-rels-val:= map:get($new-slide-map,ppt:uri-ppt-rels(()))

        (:check for uri or node val in map:)
        let $to-pres-rels := if($to-pres-rels-val instance of xs:string) then fn:doc($to-pres-rels-val)/node() else $to-pres-rels-val                         

        let $pres-rels-no-hm := ppt:remove-hm-from-pres-rels($to-pres-rels)

(:debug:) (:return $pres-rels-no-hm:)

        let $final-pres-rels := ppt:ppt-rels-insert-slide($pres-rels-no-hm, $insert-idx)

        (:update [Content_Types].xml :) 
        let $c-types-val := map:get($new-slide-map,ppt:uri-content-types(()))
        let $c-types := if($c-types-val instance of xs:string) then fn:doc($c-types-val)/node() else $c-types-val
        let $final-ctypes := ppt:update-content-types($c-types, $insert-idx, $sld-rels-img-types, $theme-ids)

        (:update presentation.xml :)
        let $pres-xml-val := map:get($new-slide-map,ppt:uri-ppt-presentation(()))
        let $pres-xml := if($pres-xml-val instance of xs:string) then fn:doc($pres-xml-val) else $pres-xml-val

        let $final-pres := ppt:update-presentation-xml($pres-xml,$final-pres-rels, $from-pres, $from-idx, $insert-idx)

        (:add 3 updates above to map:)
        let $mapupd1 := map:put( $new-slide-map, ppt:uri-ppt-presentation(()), $final-pres)
	let $mapupd2 := map:put( $new-slide-map, ppt:uri-ppt-rels(()), $final-pres-rels)
	let $mapupd3 := map:put( $new-slide-map, ppt:uri-content-types(()), $final-ctypes)

        return  $new-slide-map
};
 
(: can we improve params? :)
declare function ppt:insert-slide(
   $to-pkg-map as map:map?,
   $from-pres as xs:string+, 
   $from-idx as xs:integer+, 
   $insert-idx as xs:integer
) as map:map
{
 	let $return := 
   		if(ppt:validate-slide-indexes-map($to-pkg-map, $from-pres, $from-idx, $insert-idx)) then
      			if(ppt:validate-list-length-equal($from-pres, $from-idx)) then
            			for $from at $idx in $from-pres
            			return ppt:merge-slide-util($to-pkg-map,$from,$from-idx[$idx], $insert-idx+($idx - 1))
      			else
        			ppt:list-length-error()
   		else
        		ppt:slide-index-error()
   	return $to-pkg-map 
};

declare function ppt:package-map-zip(
   $map as map:map*
) as binary()
{
   
	let $parts := 
              for $m in $map
              let $keys := map:keys($m)
              return $keys

	let $finaldocs := 
              for $p in $parts
              let $val := map:get($map, $p)
              return if($val instance of xs:string) then fn:doc($val) else $val

	let $manifest := 
	<parts xmlns="xdmp:zip"> 
   	{
    		for $i in $parts
    		let $part :=  <part>{$i}</part>
    		return $part
   	}
        </parts>

	return xdmp:zip-create($manifest, $finaldocs)
};

declare function ppt:package-map(
   $src-dir as xs:string
) as map:map
{
  	let $doc-map := map:map()
	let $t-uris := ppt:directory-uris($src-dir)
        let $upd := for $t in $t-uris
                    return map:put($doc-map,fn:substring-after($t,$src-dir), $t)
        return $doc-map 
};

