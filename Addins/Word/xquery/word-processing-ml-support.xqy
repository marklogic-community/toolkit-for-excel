xquery version "1.0-ml";
(: Copyright 2008 Mark Logic Corporation

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
module namespace ooxml = "http://marklogic.com/openxml";

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

import module "http://marklogic.com/openxml" at "/MarkLogic/openxml/package.xqy";

declare function ooxml:create-paragraph($para as xs:string) as element(w:p)
{
  element w:p{ element w:r { element w:t {$para}}}
};

(: BEGIN REMOVE w:p PROPERTIES =============================================================== :)
declare function ooxml:passthru-para($x as node()) as node()*
{
   for $i in $x/node() return ooxml:dispatch-paragraph-to-clean($i)
};

declare function ooxml:dispatch-paragraph-to-clean($x as node()) as node()?
{

      typeswitch($x)
       case text() return $x
       case document-node() return document {$x/@*,ooxml:passthru-para($x)}
       case element(w:pPr) return ()
       case element(w:rPr) return () 
       case element() return  element{fn:name($x)} {$x/@*,passthru-para($x)}
       default return $x

};

declare function ooxml:remove-paragraph-styles($paragraph as element()) as element()
{
    ooxml:dispatch-paragraph-to-clean($paragraph)
};

(: END REMOVE w:p PROPERTIES ================================================================= :)
declare function ooxml:get-paragraph-styles($paragraph as element(w:p)*) as element(w:pPr)*
{
   $paragraph//w:pPr
};

declare function ooxml:get-run-styles($paragraph as element(w:p)*) as element(w:rPr)*
{
   $paragraph//w:rPr
};

declare function ooxml:get-paragraph-style-id($pstyle as element (w:pPr)) as xs:string?
{
   let $styles := $pstyle//w:pStyle/@w:val
   return $styles 
};

declare function ooxml:get-run-style-id($rstyle as element (w:rPr)) as xs:string?
{
   let $styles := $rstyle//w:rStyle/@w:val
   return $styles 
};

declare function ooxml:get-style-definition($styleid as xs:string, $styles as element(w:styles) ) as element(w:style)?
{
   for $id in $styleid 
   return $styles//w:style[@w:styleId=$id]
};

declare function ooxml:replace-style-definition($newstyle as element(w:style), $styles as element(w:styles)) as element(w:styles)
{
                 element w:styles { $styles/@*,
                     $styles/* except $styles//w:style[@w:styleId=$newstyle/@w:styleId],
                     $newstyle }
};



(: BEGIN SET PARAGRAPH STYLES ================================================================  :)



declare function ooxml:set-paragraph-styles-passthru($x as node()*, $props as element()?, $type as xs:string) as node()*
{
       for $i in $x/node() return ooxml:set-paragraph-styles-dispatch($i, $props, $type)
};

declare function ooxml:set-paragraph-styles-dispatch($wp as node()*, $props as element()?, $type as xs:string ) as node()*
{
       typeswitch ($wp)
         case text() return $wp
         case document-node() return document {$wp/@*,ooxml:set-paragraph-styles-passthru($wp, $props, $type)}

         case element(w:p) return if($type eq "wp") then
                                       ooxml:add-paragraph-properties($wp, $props, $type)
                                  else 
                                       element{fn:node-name($wp)} {$wp/@*,ooxml:set-paragraph-styles-passthru($wp, $props, $type)}
         case element(w:r) return if($type eq "wr") then
                                       ooxml:add-run-style-properties($wp, $props) 
                                  else 
                                       element{fn:node-name($wp)} {$wp/@*,ooxml:set-paragraph-styles-passthru($wp, $props, $type)}
         case element() return  element{fn:node-name($wp)} {$wp/@*,ooxml:set-paragraph-styles-passthru($wp, $props, $type)}
         default return $wp
          

};

declare function ooxml:add-run-style-properties($wr as node(),$runprops as element(w:rPr)? ) as node()*
{
       element w:r{ $wr/@*, $runprops, $wr/* except $wr/w:rPr }
};

declare function ooxml:add-paragraph-properties($wp as node()*, $paraprops as element(w:pPr)?, $type as xs:string) as node()*
{
        element w:p{ $wp/@*, $paraprops, ooxml:set-paragraph-styles-passthru($wp/* except $wp/w:pPr, $paraprops, $type) }
};

declare function ooxml:replace-paragraph-styles($block as element(), $wpProps as element(w:pPr)?) as element()
{
     ooxml:set-paragraph-styles-dispatch($block,$wpProps,"wp")
};

declare function ooxml:replace-run-styles($block as element(), $wrProps as element(w:rPr)?) as element()
{
     ooxml:set-paragraph-styles-dispatch($block,$wrProps,"wr")
};


(: END SET PARAGRAPH STYLES ==================================================================== :)

declare function ooxml:custom-xml($content as element(), $tag as xs:string) as element(w:customXml)?
{
  typeswitch($content)
   case element(w:p) return  element w:customXml{attribute w:element{$tag}, $content}
   case element(w:r) return  element w:customXml{attribute w:element{$tag}, $content}
   case element(w:customXml) return  element w:customXml{attribute w:element{$tag}, $content}
   case element(w:sdt) return  element w:customXml{attribute w:element{$tag}, $content}
   case element(w:tbl) return  element w:customXml{attribute w:element{$tag}, $content}
   case element(w:tr) return  element w:customXml{attribute w:element{$tag}, $content}
   case element(w:tc) return  element w:customXml{attribute w:element{$tag}, $content}
   case element(w:hyperlink) return  element w:customXml{attribute w:element{$tag}, $content}
   case element(w:fldSimple) return  element w:customXml{attribute w:element{$tag}, $content}
   case element(w:fldChar) return  element w:customXml{attribute w:element{$tag}, $content}
  default return ()
};

(: BEGIN SET CUSTOM XML TAG ==================================================================== :)
declare function ooxml:set-custom-xml-passthru($x as node()*, $oldtag as xs:string, $newtag as xs:string) as node()*
{
       for $i in $x/node() return ooxml:set-custom-xml-dispatch($i, $oldtag, $newtag)
};

declare function ooxml:set-custom-xml-dispatch($block as node()*, $oldtag as xs:string, $newtag as xs:string) as node()*
{
       typeswitch ($block)
       case text() return $block
       case document-node() return document {$block/@*,ooxml:set-custom-xml-passthru($block, $oldtag, $newtag)}
       case element(w:customXml) return ooxml:set-custom-element-value($block, $oldtag, $newtag) 
       case element() return  element{fn:node-name($block)} {$block/@*,ooxml:set-custom-xml-passthru($block, $oldtag, $newtag)}
       default return $block
};

declare function ooxml:set-custom-element-value($block as node()*, $oldtag as xs:string, $newtag as xs:string) as node()*
{
   let $value := $block/@w:element
   let $cxml := if($value eq $oldtag) then
                      element w:customXml {attribute w:element{$newtag}, $block/@* except $block/@w:element, ooxml:set-custom-xml-passthru($block, $oldtag, $newtag)}
                   else
                      element{fn:node-name($block)} {$block/@*,ooxml:set-custom-xml-passthru($block, $oldtag, $newtag)} 
   return $cxml
};

declare function ooxml:replace-custom-xml-element($content as element(), $oldtag as xs:string, $newtag as xs:string) as element()
{ 
    let $newblock := ooxml:set-custom-xml-dispatch($content, $oldtag, $newtag) 
    return $newblock
};
(: END SET CUSTOM XML TAG ====================================================================== :)

declare function ooxml:get-custom-xml-ancestor($doc as element()) as element()?
{

    let $uri := "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    let $tmp := $doc
    let $ancestor := 
         if($tmp/parent::w:sdtContent) then ooxml:get-custom-xml-ancestor($tmp/../..) 
         else if($tmp/parent::w:customXml) then ooxml:get-custom-xml-ancestor($tmp/..)
         else  $tmp
    let $nodename := fn:node-name($ancestor)
    let $final :=  if($nodename eq fn:QName($uri,"customXml") or $nodename eq fn:QName($uri,"sdt")) then $ancestor else ()
    return $final
 
};

(: BEGIN SIMPLE SEARCH ================================================================================ :)

declare function ooxml:paragraph-search($query as cts:query) as node()*
{
    let $doc := cts:search(//w:p ,$query)
    return $doc
};

declare function ooxml:paragraph-search($query as cts:query, $begin as xs:integer, $end as xs:integer) as node()*
{
    let $doc := cts:search(//w:p ,$query)[$begin to $end]
    return $doc
};

declare function ooxml:custom-search-all($query as cts:query, $begin as xs:integer, $end as xs:integer) as node()*
{
    let $sdt := cts:search( //(w:sdt | w:customXml | w:p ), ($query))[$begin to $end]
    return $sdt
};

(: END SIMPLE SEARCH ================================================================================== :)

(: BEGIN w:customXml HIGHLIGHT ================================================================= :)

declare function ooxml:passthru-chlt($x as node()*) as node()*
{
  for $i in $x/node() return ooxml:dispatch-chlt($i)
};

declare function ooxml:map($props as node()*, $x as node()*) as node()*
{

  for $child in $x return
   typeswitch ($child)
    case text() return ooxml:makerun($child, $props)
    case element(w:customXml) return element{fn:name($child)} {$child/@*, $child/w:customXmlPr, <w:r>{$props,$child/w:r/child::*}</w:r>}
    case element() return element{fn:name($x)} {$x/@*,ooxml:passthru-chlt($x)}
    default return $x
};


declare function ooxml:dispatch-chlt($x as node()*) as node()*
{
   typeswitch ($x)
    case document-node() return ooxml:passthru-chlt($x)
    case text() return $x
    case element(w:r) return (if(fn:exists($x//child::*//w:p)) then ooxml:passthru-chlt($x) 
                              else ooxml:map((if(fn:empty($x/w:rPr/node())) then () else <w:rPr>{$x/w:rPr/node()}</w:rPr>), $x/w:t/node()))
    case element() return  element{fn:name($x)} {$x/@*,ooxml:passthru-chlt($x)} 
    default return $x
};

declare function ooxml:makerun($x as text(), $runProps as element(w:rPr)?) as element(w:r)
{ 
    <w:r>{$runProps}<w:t xml:space="preserve">{$x}</w:t></w:r>
};

declare function ooxml:custom-xml-highlight-exec($orig as node()*, $query as cts:query, $tagname as xs:string, $attrs as xs:string*, $vals as xs:string*) as node()*
{    let $tmpdoc := <temp>{$orig}</temp>
     let $highlightedbody := cts:highlight($tmpdoc, $query, 
                               <w:customXml w:element="{$tagname}">
                                { if(fn:count($attrs) gt 0 )
                                  then
                                   <w:customXmlPr>
                                    {
                                     for $attr at $d in $attrs 
                                      return <w:attr w:name ={$attr}  w:val={$vals[$d]} />
                                    }
                                    </w:customXmlPr>
                                   else ()
                                }    
                                    <w:r><w:t>{$cts:text}</w:t></w:r>
                               </w:customXml>)
     let $newdocument := ooxml:dispatch-chlt($highlightedbody)
     return $newdocument/*
};

declare function ooxml:custom-xml-highlight-exec($orig as node()*, $query as cts:query, $tagname as xs:string) as node()*
{    let $tmpdoc := <temp>{$orig}</temp>
     let $highlightedbody := cts:highlight($tmpdoc, $query, 
                               <w:customXml w:element="{$tagname}">
                                    <w:r><w:t>{$cts:text}</w:t></w:r>
                               </w:customXml>)
     let $newdocument := ooxml:dispatch-chlt($highlightedbody)
     return $newdocument/* 
};


declare function ooxml:custom-xml-highlight($nodes as node()*, $highlight-term as cts:query, $tag-name as xs:string,  $attributes as xs:string*, $values as xs:string*) as  node()*
{
   let $return := if(ooxml:validate-list-length-equal-strings($attributes,$values)) then 
      ooxml:custom-xml-highlight-exec($nodes,$highlight-term,$tag-name, $attributes, $values)
   else ooxml:list-length-error()
   return $return
};

declare function ooxml:custom-xml-highlight($nodes as node()*, $highlight-term as cts:query, $tag-name as xs:string) as  node()*
{
      ooxml:custom-xml-highlight-exec($nodes,$highlight-term,$tag-name)
};

(: added Entity Hightlight :)
declare function ooxml:custom-xml-entity-hlt($nodes as node()*) as node()*
{
   let $enriched := cts:entity-highlight($nodes, element w:customXml{attribute w:element{fn:replace($cts:entity-type, ":", "-")}, <w:r><w:t>{$cts:text}</w:t></w:r> } )
   let $final := ooxml:dispatch-chlt($enriched)
   return $final
  
};

(: END w:customXml HIGHLIGHT =================================================================== :)

(: added OPC Package serialization support :)

declare function ooxml:formatbinary($s as xs:string*) as xs:string*
{
 if(fn:string-length($s) > 0) then
     let $firstpart := fn:concat(fn:substring($s,1,76))
     let $tail := fn:substring-after($s,$firstpart)
     return ($firstpart,ooxml:formatbinary($tail))
                  else
             ()
  

  (: let $x := "PETE"
     let $y := "O"
     return ($x, $y)
  :)
};

declare function ooxml:get-part-content-type($uri as xs:string) as xs:string?
{
   if(fn:ends-with($uri,".rels"))
   then 
        "application/vnd.openxmlformats-package.relationships+xml"
   else if(fn:ends-with($uri,"glossary/document.xml"))
   then
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml"
   else if(fn:ends-with($uri,"document.xml"))
   then
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml" 
   else if(fn:matches($uri, "theme\d+\.xml"))
   then 
      "application/vnd.openxmlformats-officedocument.theme+xml"
   else if(fn:ends-with($uri,"word/numbering.xml"))
   then 
      "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"
   (: else if(fn:ends-with($uri,"word/settings.xml")):)
   else if(fn:ends-with($uri,"settings.xml"))
   then 
      "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"
   (:else if(fn:ends-with($uri,"word/styles.xml")):)
   else if(fn:ends-with($uri,"styles.xml"))
   then 
      "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"
   (: else if(fn:ends-with($uri,"word/webSettings.xml")) :)
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
   else if(fn:matches($uri,"customXml/itemProps\d+\.xml")) then
      "application/vnd.openxmlformats-officedocument.customXmlProperties+xml"
   else if(fn:matches($uri,"customXml/item\d+\.xml")) then
      "application/xml"
   else
       ()
    
};

declare function ooxml:get-part-attributes($uri as xs:string) as node()*
{
  (:not sure if this is needed, for serverside generated docx, path comes through as \path\name , if path has /mixed\separators\for\path/example, it chokes when opening in word :)
  let $cleanuri := fn:replace($uri,"\\","/")
  let $name := attribute pkg:name{$cleanuri}
  let $contenttype := attribute pkg:contentType{ooxml:get-part-content-type($cleanuri)}
  let $padding := if(fn:ends-with($cleanuri,".rels")) then

                     if(fn:starts-with($cleanuri,"/word/glossary")) then
                         ()
                    
                     else if(fn:starts-with($cleanuri,"/_rels")) then
                      attribute pkg:padding{ "512" }
                     else    
                      attribute pkg:padding{ "256" }
                  else
                     ()
  let $compression := if(fn:ends-with($cleanuri,"jpeg")) then 
                         attribute pkg:compression { "store" } 
                      else ()
  
  return ($name, $contenttype, $padding, $compression)
};

declare function ooxml:get-package-part($directory as xs:string, $uri as xs:string) as node()?
{
  let $fulluri := $uri
  let $docuri := fn:concat("/",fn:substring-after($fulluri,$directory))
  let $data := fn:doc($fulluri)

  let $part := if(fn:empty($data) or fn:ends-with($fulluri,"[Content_Types].xml")) then () 
               else if(fn:ends-with($fulluri,".jpeg") or fn:ends-with($fulluri,".wmf")) then
                  let $bin :=   xs:base64Binary(xs:hexBinary($data)) cast as xs:string 
                    let $formattedbin := fn:string-join(ooxml:formatbinary($bin),"&#x9;&#xA;") 
                  (: let $formattedbin := fn:string-join(ooxml:formatbinary($bin),"\r\n") :)
                  (:let $formattedbin := mlos:formatbinary($bin):)
                  return  element pkg:part { ooxml:get-part-attributes($docuri), element pkg:binaryData { $formattedbin  }   }
                  (: element pkg:part { mlos:get-part-attributes($docuri), element pkg:binaryData {  xs:base64Binary(xs:hexBinary($data))    }   } :)
               else
                  element pkg:part { ooxml:get-part-attributes($docuri), element pkg:xmlData { $data }}
  return  $part (: <T>{$fulluri}</T>   :) 
};

declare function ooxml:make-package($directory as xs:string, $uris as xs:string*) as node()*
{
  let $package := element pkg:package { 
                            for $uri in $uris
                            let $part := ooxml:get-package-part($directory,$uri)
                            return $part }
                           
return 
(: <?mso-application progid="Word.Document"?>, :)
     $package
};

declare function ooxml:package-uris-from-directory($docuri as xs:string) as xs:string*
{

  cts:uris("","document",cts:directory-query($docuri,"infinity"))

};

declare function ooxml:package-files-only($uris as xs:string*) as xs:string*
{
                  for $uri in $uris
                  let $u := if(fn:ends-with($uri,"/")) then () else $uri
                  return $u
};
(: end OPC Package serialization :)

