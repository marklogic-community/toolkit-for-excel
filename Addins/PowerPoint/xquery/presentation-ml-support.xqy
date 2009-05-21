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
declare namespace r="http://schemas.openxmlformats.org/officeDocument/2006/relationships";
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
 if(fn:string-length($s) > 0) then
     let $firstpart := fn:concat(fn:substring($s,1,76))
     let $tail := fn:substring-after($s,$firstpart)
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

declare function ppt:package-files-only($uris as xs:string*) as xs:string*
{
                  for $uri in $uris
                  let $u := if(fn:ends-with($uri,"/")) then () else $uri
                  return $u
};
