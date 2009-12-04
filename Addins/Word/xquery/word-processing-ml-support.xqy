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

declare namespace a="http://schemas.openxmlformats.org/drawingml/2006/main";
declare namespace w="http://schemas.openxmlformats.org/wordprocessingml/2006/main";
declare namespace v="urn:schemas-microsoft-com:vml";
declare namespace ve="http://schemas.openxmlformats.org/markup-compatibility/2006";
declare namespace o="urn:schemas-microsoft-com:office:office";
declare namespace m="http://schemas.openxmlformats.org/officeDocument/2006/math";
declare namespace wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
declare namespace w10="urn:schemas-microsoft-com:office:word";
declare namespace wne="http://schemas.microsoft.com/office/word/2006/wordml";
declare namespace pkg="http://schemas.microsoft.com/office/2006/xmlPackage";
declare namespace pic="http://schemas.openxmlformats.org/drawingml/2006/picture";
declare namespace r="http://schemas.openxmlformats.org/officeDocument/2006/relationships";
declare namespace pr = "http://schemas.openxmlformats.org/package/2006/relationships";
declare namespace types = "http://schemas.openxmlformats.org/package/2006/content-types";
declare namespace zip   = "xdmp:zip";

import module "http://marklogic.com/openxml" at "/MarkLogic/openxml/package.xqy";

declare variable $ooxml:TYPES := "http://schemas.openxmlformats.org/package/2006/content-types";
declare variable $ooxml:PKG-RELATIONSHIPS := "http://schemas.openxmlformats.org/package/2006/relationships";
declare variable $ooxml:DOC-RELATIONSHIPS := "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
declare variable $ooxml:DRAWINGML := "http://schemas.openxmlformats.org/drawingml/2006/main";
declare variable $ooxml:WORDPROCESSINGML := "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
declare variable $ooxml:VML := "urn:schemas-microsoft-com:vml";
declare variable $ooxml:COMPATABILITY := "http://schemas.openxmlformats.org/markup-compatibility/2006";
declare variable $ooxml:OFFICE := "urn:schemas-microsoft-com:office:office";
declare variable $ooxml:MATH := "http://schemas.openxmlformats.org/officeDocument/2006/math";
declare variable $ooxml:WORDPROCESSING-DRAWING := "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
declare variable $ooxml:WORD := "urn:schemas-microsoft-com:office:word";
declare variable $ooxml:WORDML := "http://schemas.microsoft.com/office/word/2006/wordml";
declare variable $ooxml:PACKAGE := "http://schemas.microsoft.com/office/2006/xmlPackage";
declare variable $ooxml:PICTURE := "http://schemas.openxmlformats.org/drawingml/2006/picture";
declare variable $ooxml:EXT-PROPERTIES := "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
declare variable $ooxml:CORE-PROPERTIES := "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
declare variable $ooxml:CUSTOM-PROPERTIES := "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";
declare variable $ooxml:CUSTOM-XML-PROPS := "http://schemas.openxmlformats.org/officeDocument/2006/customXml";

declare function ooxml:error($message as xs:string)
{
    fn:error(xs:QName("DOCX-ERROR"),$message)
};

declare function ooxml:get-mimetype(
  $filename as xs:string
) as xs:string?
{
    xdmp:uri-content-type($filename)	
};

declare function ooxml:directory-uris(
  $directory as xs:string
) as xs:string*
{
    ooxml:directory-uris($directory,"infinity")
};

declare function ooxml:directory-uris(
  $directory as xs:string, 
  $depth as xs:string
) as xs:string*
{
    cts:uris("","document",cts:directory-query($directory,$depth))
};

declare function ooxml:create-paragraph(
  $para as xs:string
) as element(w:p)
{
    element w:p{ element w:r { element w:t {$para}}}
};

(: BEGIN REMOVE w:p PROPERTIES ==============================================  :)
declare function ooxml:passthru-para(
  $x as node()
) as node()*
{
    for $i in $x/node() return ooxml:dispatch-paragraph-to-clean($i)
};

declare function ooxml:dispatch-paragraph-to-clean(
  $x as node()
) as node()?
{
    typeswitch($x)
      case text() return $x
      case document-node() return document {$x/@*,ooxml:passthru-para($x)}
      case element(w:pPr) return ()
      case element(w:rPr) return () 
      case element() return  element{fn:name($x)} {$x/@*,passthru-para($x)}
      default return $x
};

declare function ooxml:remove-paragraph-styles(
  $paragraph as element()
) as element()
{
    ooxml:dispatch-paragraph-to-clean($paragraph)
};
(: END REMOVE w:p PROPERTIES ================================================  :)

declare function ooxml:get-paragraph-styles(
  $paragraph as element(w:p)*
) as element(w:pPr)*
{
     $paragraph//w:pPr
};

declare function ooxml:get-run-styles(
  $paragraph as element(w:p)*
) as element(w:rPr)*
{
    $paragraph//w:rPr
};

declare function ooxml:get-paragraph-style-id(
  $pstyle as element (w:pPr)
) as xs:string?
{
    let $styles := $pstyle//w:pStyle/@w:val
    return $styles 
};

declare function ooxml:get-run-style-id(
  $rstyle as element (w:rPr)
) as xs:string?
{
    let $styles := $rstyle//w:rStyle/@w:val
    return $styles 
};

declare function ooxml:get-style-definition(
  $styleid as xs:string, 
  $styles as element(w:styles)
) as element(w:style)?
{
    for $id in $styleid 
    return $styles//w:style[@w:styleId=$id]
};

declare function ooxml:replace-style-definition(
  $newstyle as element(w:style), 
  $styles as element(w:styles)
) as element(w:styles)
{
    element w:styles { $styles/@*,
                       $styles/* except $styles//w:style[@w:styleId=$newstyle/@w:styleId],
                       $newstyle }
};

(: BEGIN SET PARAGRAPH STYLES ===============================================  :)
declare function ooxml:set-paragraph-styles-passthru(
  $x as node()*, 
  $props as element()?, 
  $type as xs:string
) as node()*
{
    for $i in $x/node() return ooxml:set-paragraph-styles-dispatch($i, $props, $type)
};

declare function ooxml:set-paragraph-styles-dispatch(
  $wp as node()*, 
  $props as element()?, 
  $type as xs:string 
) as node()*
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

declare function ooxml:add-run-style-properties(
  $wr as node(),
  $runprops as element(w:rPr)?
) as node()*
{
    element w:r{ $wr/@*, $runprops, $wr/* except $wr/w:rPr }
};

declare function ooxml:add-paragraph-properties(
  $wp as node()*, 
  $paraprops as element(w:pPr)?, 
  $type as xs:string
) as node()*
{
    element w:p{ $wp/@*, $paraprops, ooxml:set-paragraph-styles-passthru($wp/* except $wp/w:pPr, $paraprops, $type) }
};

declare function ooxml:replace-paragraph-styles(
  $block as element(), 
  $wpProps as element(w:pPr)?
) as element()
{
    ooxml:set-paragraph-styles-dispatch($block,$wpProps,"wp")
};

declare function ooxml:replace-run-styles(
  $block as element(), 
  $wrProps as element(w:rPr)?
) as element()
{
    ooxml:set-paragraph-styles-dispatch($block,$wrProps,"wr")
};
(: END SET PARAGRAPH STYLES =================================================  :)

declare function ooxml:custom-xml(
  $content as element(), 
  $tag as xs:string
) as element(w:customXml)?
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

(: BEGIN SET CUSTOM XML TAG =================================================  :)
declare function ooxml:set-custom-xml-passthru(
  $x as node()*, 
  $oldtag as xs:string, 
  $newtag as xs:string
) as node()*
{
    for $i in $x/node() return ooxml:set-custom-xml-dispatch($i, $oldtag, $newtag)
};

declare function ooxml:set-custom-xml-dispatch(
  $block as node()*, 
  $oldtag as xs:string, 
  $newtag as xs:string
) as node()*
{
       typeswitch ($block)
       case text() return $block
       case document-node() return document {$block/@*,ooxml:set-custom-xml-passthru($block, $oldtag, $newtag)}
       case element(w:customXml) return ooxml:set-custom-element-value($block, $oldtag, $newtag) 
       case element() return  element{fn:node-name($block)} {$block/@*,ooxml:set-custom-xml-passthru($block, $oldtag, $newtag)}
       default return $block
};

declare function ooxml:set-custom-element-value(
  $block as node()*, 
  $oldtag as xs:string, 
  $newtag as xs:string
) as node()*
{
    let $value := $block/@w:element
    let $cxml := if($value eq $oldtag) then
                     element w:customXml {attribute w:element{$newtag}, $block/@* except $block/@w:element, ooxml:set-custom-xml-passthru($block, $oldtag, $newtag)}
                 else
                     element{fn:node-name($block)} {$block/@*,ooxml:set-custom-xml-passthru($block, $oldtag, $newtag)} 
    return $cxml
};

declare function ooxml:replace-custom-xml-element(
  $content as element(), 
  $oldtag as xs:string, 
  $newtag as xs:string
) as element()
{ 
    let $newblock := ooxml:set-custom-xml-dispatch($content, $oldtag, $newtag) 
    return $newblock
};
(: END SET CUSTOM XML TAG ===================================================  :)

declare function ooxml:get-custom-xml-ancestor(
  $doc as element()
) as element()?
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

(: BEGIN SIMPLE SEARCH ======================================================  :)
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
(: END SIMPLE SEARCH ========================================================  :)

(: BEGIN w:customXml HIGHLIGHT ==============================================  :)
declare function ooxml:passthru-chlt(
  $x as node()*
) as node()*
{
    for $i in $x/node() return ooxml:dispatch-chlt($i)
};

declare function ooxml:map(
  $props as node()*, 
  $x as node()*
) as node()*
{

    for $child in $x return
     typeswitch ($child)
       case text() return ooxml:makerun($child, $props)
       case element(w:customXml) return element{fn:name($child)} {$child/@*, $child/w:customXmlPr, <w:r>{$props,$child/w:r/child::*}</w:r>}
       case element() return element{fn:name($x)} {$x/@*,ooxml:passthru-chlt($x)}
       default return $x
};


declare function ooxml:dispatch-chlt(
  $x as node()*
) as node()*
{
    typeswitch ($x)
      case document-node() return ooxml:passthru-chlt($x)
      case text() return $x
      case element(w:r) return (if(fn:exists($x//child::*//w:p)) then ooxml:passthru-chlt($x) 
                                else ooxml:map((if(fn:empty($x/w:rPr/node())) then () else <w:rPr>{$x/w:rPr/node()}</w:rPr>), $x/w:t/node()))
      case element() return  element{fn:name($x)} {$x/@*,ooxml:passthru-chlt($x)} 
      default return $x
};

declare function ooxml:makerun(
  $x as text(), 
  $runProps as element(w:rPr)?
) as element(w:r)
{ 
    <w:r>{$runProps}<w:t xml:space="preserve">{$x}</w:t></w:r>
};

declare function ooxml:custom-xml-highlight-exec(
  $orig as node()*, 
  $query as cts:query, 
  $tagname as xs:string, 
  $attrs as xs:string*, 
  $vals as xs:string*
) as node()*
{    
    let $tmpdoc := <temp>{$orig}</temp>
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

declare function ooxml:custom-xml-highlight-exec(
  $orig as node()*, 
  $query as cts:query, 
  $tagname as xs:string
) as node()*
{
    let $tmpdoc := <temp>{$orig}</temp>
    let $highlightedbody := cts:highlight($tmpdoc, $query, 
                               <w:customXml w:element="{$tagname}">
                                    <w:r><w:t>{$cts:text}</w:t></w:r>
                               </w:customXml>)
    let $newdocument := ooxml:dispatch-chlt($highlightedbody)
    return $newdocument/* 
};


declare function ooxml:custom-xml-highlight(
  $nodes as node()*, 
  $highlight-term as cts:query, 
  $tag-name as xs:string,  
  $attributes as xs:string*, 
  $values as xs:string*
) as  node()*
{
    let $return := if(ooxml:validate-list-length-equal-strings($attributes,$values)) then 
                    ooxml:custom-xml-highlight-exec($nodes,$highlight-term,$tag-name, $attributes, $values)
                   else ooxml:list-length-error()
    return $return
};

declare function ooxml:custom-xml-highlight(
  $nodes as node()*, 
  $highlight-term as cts:query, 
  $tag-name as xs:string
) as  node()*
{
    ooxml:custom-xml-highlight-exec($nodes,$highlight-term,$tag-name)
};

declare function ooxml:custom-xml-entity-highlight( (: normalized text ? :)
  $nodes as node()*
) as node()*
{
    let $enriched := 
          cts:entity-highlight($nodes, 
                               element w:customXml{attribute w:element{fn:replace($cts:entity-type, ":", "-")}, 
                               <w:r><w:t>{$cts:text}</w:t></w:r> })
    let $final := ooxml:dispatch-chlt($enriched)
    return $final
  
};
(: END w:customXml HIGHLIGHT ================================================  :)

(: added OPC Package serialization support :)
declare function ooxml:format-binary(
$binstring as xs:string
)as xs:string*
{
    for $i in 0 to (fn:string-length($binstring) idiv 76)
    let $start := ($i * 76)
    return fn:substring($binstring,$start,76) 
};

declare function ooxml:get-part-content-type(
  $node as node()
) as xs:string?
{
    if(fn:node-name($node) eq fn:QName($ooxml:PKG-RELATIONSHIPS, "Relationships"))
    then 
        "application/vnd.openxmlformats-package.relationships+xml"
    else if(fn:node-name($node) eq fn:QName($ooxml:WORDPROCESSINGML, "document")) 
    then
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml" 
    else if(fn:node-name($node) eq fn:QName($ooxml:DRAWINGML, "theme")) 
    then
        "application/vnd.openxmlformats-officedocument.theme+xml"
    else if(fn:node-name($node) eq fn:QName($ooxml:WORDPROCESSINGML, "numbering")) 
    then
        "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"
    else if(fn:node-name($node) eq fn:QName($ooxml:WORDPROCESSINGML, "settings")) 
    then
        "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"
    else if(fn:node-name($node) eq fn:QName($ooxml:WORDPROCESSINGML, "styles"))
    then 
        "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"
    else if(fn:node-name($node) eq fn:QName($ooxml:WORDPROCESSINGML, "fonts"))
    then
        "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"
    else if(fn:node-name($node) eq fn:QName($ooxml:WORDPROCESSINGML, "webSettings"))
    then
        "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml"
    else if(fn:node-name($node) eq fn:QName($ooxml:WORDPROCESSINGML, "ftr"))
    then 
        "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"
    else if(fn:node-name($node) eq fn:QName($ooxml:WORDPROCESSINGML, "hdr"))
    then
        "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
    else if(fn:node-name($node) eq fn:QName($ooxml:WORDPROCESSINGML, "endnotes"))
    then
        "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"
    else if(fn:node-name($node) eq fn:QName($ooxml:WORDPROCESSINGML, "footnotes"))
    then
        "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"
    else if(fn:node-name($node) eq fn:QName($ooxml:EXT-PROPERTIES, "Properties"))
    then
        "application/vnd.openxmlformats-officedocument.extended-properties+xml"
    else if(fn:node-name($node) eq fn:QName($ooxml:CORE-PROPERTIES, "coreProperties"))
    then
        "application/vnd.openxmlformats-package.core-properties+xml"
    else if(fn:node-name($node) eq fn:QName($ooxml:CUSTOM-PROPERTIES, "Properties"))
    then
        "application/vnd.openxmlformats-officedocument.custom-properties+xml"
    else if(fn:node-name($node) eq fn:QName($ooxml:CUSTOM-XML-PROPS, "datastoreItem"))
    then 
        "application/vnd.openxmlformats-officedocument.customXmlProperties+xml"
    (: else if(fn:matches($uri,"customXml/item\d+\.xml")) then :)
    else 
        "application/xml"

   (:  need to account for glossary
    else if(fn:ends-with($uri,"glossary/document.xml"))
    then
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml"
   :)
};


declare function get-image-part-content-type(
  $uri as xs:string
)as xs:string?
{
    if(fn:ends-with(fn:upper-case($uri),"JPEG")) 
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
    else ()
};

declare function ooxml:get-part-attributes(
  $node as node()
) as node()*
{
    let $uri := fn:substring-after(fn:base-uri($node), "docx_parts")
    let $name := attribute pkg:name{$uri}

    let $contenttype := if (xdmp:node-kind($node) eq "binary") then
                            attribute pkg:contentType{ooxml:get-image-part-content-type($uri)}
                        else
                             attribute pkg:contentType{ooxml:get-part-content-type($node)} 

    let $padding := if(fn:ends-with($uri,".rels")) then

                      if(fn:starts-with($uri,"/word/glossary")) then
                          ()
                    
                      else if(fn:starts-with($uri,"/_rels")) then
                       attribute pkg:padding{ "512" }
                      else    
                       attribute pkg:padding{ "256" }
                    else
                      ()

    (:not required for .WMF:)
    let $compression := if(fn:ends-with(fn:upper-case($uri),"JPEG") or 
                           fn:ends-with(fn:upper-case($uri),"PNG") or
                           fn:ends-with(fn:upper-case($uri),"GIF")) then 
                             attribute pkg:compression { "store" } 
                        else ()
 
  
    return ($name, $contenttype, $padding, $compression)
};

declare function ooxml:get-package-part(
  $node as node()
) as node()?
{
    let $part := if(fn:empty($node) or 
                   (fn:node-name($node) eq fn:QName($ooxml:TYPES, "Types"))) then () 
                 else if(xdmp:node-kind($node) eq "binary") then 
                       let $bin :=   xs:base64Binary(xs:hexBinary($node)) cast as xs:string  
                       let $formattedbin := fn:string-join(ooxml:format-binary($bin),"&#x9;&#xA;") 
                       return element pkg:part { ooxml:get-part-attributes($node), element pkg:binaryData { $formattedbin  } } 
                 else 
                       element pkg:part { ooxml:get-part-attributes($node), element pkg:xmlData { $node }}
  return  $part 
};

(: two functions, one function, give the directory, call the other one, which takes sequence of nodes :)
declare function ooxml:package(
  $directory as xs:string
) as element(pkg:package)*
{

    let $uris := ooxml:directory-uris($directory)
    let $validuris := ooxml:package-files-only($uris)
    return
            element pkg:package { 
                    for $uri in $validuris
                    let $part := ooxml:get-package-part(fn:doc($uri)/node())
                    return $part 
                                }      
};

(: processing instructions generated when Word or PPT 'Save As' XML:)
(: not currently required for Office to open file :)
(: <?mso-application progid="Word.Document"?>, $package :)
(: <?mso-application progid="PowerPoint.Show"?> :)

declare function ooxml:package-files-only(
  $uris as xs:string*
) as xs:string*
{
    $uris[fn:not(fn:ends-with(.,"/"))] 
};

(: end OPC Package serialization :)


(:BEGIN new functions for server side document creation :)
declare function ooxml:package-rels(  (: default-package-rels :)
) as element(pr:Relationships)
{   
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="/word/document.xml" />
    </Relationships>
     
};

declare function ooxml:content-types( (:default-content-types, simple-content-types for now,with eye to richer function in the future, look at ppt again :)
  $default as xs:boolean
) as element(types:Types)
{
    if($default) then
       <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Default Extension="xml" ContentType="application/xml"/>
          <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
          <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
          <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
          <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
          <Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
          <Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>
       </Types> 
    else
       <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
	  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />
	  <Default Extension="xml" ContentType="application/xml" />
	  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml" />
       </Types>
};

declare function ooxml:font-table(
) as element(w:fonts)
{
<w:fonts xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:font w:name="Symbol">
    <w:panose1 w:val="05050102010706020507"/>
    <w:charset w:val="02"/>
    <w:family w:val="roman"/>
    <w:pitch w:val="variable"/>
    <w:sig w:usb0="00000000" w:usb1="10000000" w:usb2="00000000" w:usb3="00000000" w:csb0="80000000" w:csb1="00000000"/>
  </w:font>
  <w:font w:name="Times New Roman">
    <w:panose1 w:val="02020603050405020304"/>
    <w:charset w:val="00"/>
    <w:family w:val="roman"/>
    <w:pitch w:val="variable"/>
    <w:sig w:usb0="20002A87" w:usb1="80000000" w:usb2="00000008" w:usb3="00000000" w:csb0="000001FF" w:csb1="00000000"/>
  </w:font>
  <w:font w:name="Courier New">
    <w:panose1 w:val="02070309020205020404"/>
    <w:charset w:val="00"/>
    <w:family w:val="modern"/>
    <w:pitch w:val="fixed"/>
    <w:sig w:usb0="20002A87" w:usb1="80000000" w:usb2="00000008" w:usb3="00000000" w:csb0="000001FF" w:csb1="00000000"/>
  </w:font>
  <w:font w:name="Wingdings">
    <w:panose1 w:val="05000000000000000000"/>
    <w:charset w:val="02"/>
    <w:family w:val="auto"/>
    <w:pitch w:val="variable"/>
    <w:sig w:usb0="00000000" w:usb1="10000000" w:usb2="00000000" w:usb3="00000000" w:csb0="80000000" w:csb1="00000000"/>
  </w:font>
  <w:font w:name="Calibri">
    <w:panose1 w:val="020F0502020204030204"/>
    <w:charset w:val="00"/>
    <w:family w:val="swiss"/>
    <w:pitch w:val="variable"/>
    <w:sig w:usb0="A00002EF" w:usb1="4000207B" w:usb2="00000000" w:usb3="00000000" w:csb0="0000009F" w:csb1="00000000"/>
  </w:font>
  <w:font w:name="Cambria">
    <w:panose1 w:val="02040503050406030204"/>
    <w:charset w:val="00"/>
    <w:family w:val="roman"/>
    <w:pitch w:val="variable"/>
    <w:sig w:usb0="A00002EF" w:usb1="4000004B" w:usb2="00000000" w:usb3="00000000" w:csb0="0000009F" w:csb1="00000000"/>
  </w:font>
</w:fonts>

};

declare function ooxml:numbering(
) as element(w:numbering)
{
<w:numbering xmlns:ve="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml">
  <w:abstractNum w:abstractNumId="0">
    <w:nsid w:val="024D68ED"/>
    <w:multiLevelType w:val="hybridMultilevel"/>
    <w:tmpl w:val="4A228AE4"/>
    <w:lvl w:ilvl="0" w:tplc="0409000F">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%1."/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="720" w:hanging="360"/>
      </w:pPr>
    </w:lvl>
    <w:lvl w:ilvl="1" w:tplc="04090019" w:tentative="1">
      <w:start w:val="1"/>
      <w:numFmt w:val="lowerLetter"/>
      <w:lvlText w:val="%2."/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="1440" w:hanging="360"/>
      </w:pPr>
    </w:lvl>
    <w:lvl w:ilvl="2" w:tplc="0409001B" w:tentative="1">
      <w:start w:val="1"/>
      <w:numFmt w:val="lowerRoman"/>
      <w:lvlText w:val="%3."/>
      <w:lvlJc w:val="right"/>
      <w:pPr>
        <w:ind w:left="2160" w:hanging="180"/>
      </w:pPr>
    </w:lvl>
    <w:lvl w:ilvl="3" w:tplc="0409000F" w:tentative="1">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%4."/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="2880" w:hanging="360"/>
      </w:pPr>
    </w:lvl>
    <w:lvl w:ilvl="4" w:tplc="04090019" w:tentative="1">
      <w:start w:val="1"/>
      <w:numFmt w:val="lowerLetter"/>
      <w:lvlText w:val="%5."/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="3600" w:hanging="360"/>
      </w:pPr>
    </w:lvl>
    <w:lvl w:ilvl="5" w:tplc="0409001B" w:tentative="1">
      <w:start w:val="1"/>
      <w:numFmt w:val="lowerRoman"/>
      <w:lvlText w:val="%6."/>
      <w:lvlJc w:val="right"/>
      <w:pPr>
        <w:ind w:left="4320" w:hanging="180"/>
      </w:pPr>
    </w:lvl>
    <w:lvl w:ilvl="6" w:tplc="0409000F" w:tentative="1">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%7."/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="5040" w:hanging="360"/>
      </w:pPr>
    </w:lvl>
    <w:lvl w:ilvl="7" w:tplc="04090019" w:tentative="1">
      <w:start w:val="1"/>
      <w:numFmt w:val="lowerLetter"/>
      <w:lvlText w:val="%8."/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="5760" w:hanging="360"/>
      </w:pPr>
    </w:lvl>
    <w:lvl w:ilvl="8" w:tplc="0409001B" w:tentative="1">
      <w:start w:val="1"/>
      <w:numFmt w:val="lowerRoman"/>
      <w:lvlText w:val="%9."/>
      <w:lvlJc w:val="right"/>
      <w:pPr>
        <w:ind w:left="6480" w:hanging="180"/>
      </w:pPr>
    </w:lvl>
  </w:abstractNum>
  <w:abstractNum w:abstractNumId="1">
    <w:nsid w:val="0C8C0FAC"/>
    <w:multiLevelType w:val="hybridMultilevel"/>
    <w:tmpl w:val="ECCE2F5C"/>
    <w:lvl w:ilvl="0" w:tplc="04090001">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="?"/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="720" w:hanging="360"/>
      </w:pPr>
      <w:rPr>
        <w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/>
      </w:rPr>
    </w:lvl>
    <w:lvl w:ilvl="1" w:tplc="04090003" w:tentative="1">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="o"/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="1440" w:hanging="360"/>
      </w:pPr>
      <w:rPr>
        <w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:cs="Courier New" w:hint="default"/>
      </w:rPr>
    </w:lvl>
    <w:lvl w:ilvl="2" w:tplc="04090005" w:tentative="1">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="?"/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="2160" w:hanging="360"/>
      </w:pPr>
      <w:rPr>
        <w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/>
      </w:rPr>
    </w:lvl>
    <w:lvl w:ilvl="3" w:tplc="04090001" w:tentative="1">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="?"/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="2880" w:hanging="360"/>
      </w:pPr>
      <w:rPr>
        <w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/>
      </w:rPr>
    </w:lvl>
    <w:lvl w:ilvl="4" w:tplc="04090003" w:tentative="1">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="o"/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="3600" w:hanging="360"/>
      </w:pPr>
      <w:rPr>
        <w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:cs="Courier New" w:hint="default"/>
      </w:rPr>
    </w:lvl>
    <w:lvl w:ilvl="5" w:tplc="04090005" w:tentative="1">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="?"/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="4320" w:hanging="360"/>
      </w:pPr>
      <w:rPr>
        <w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/>
      </w:rPr>
    </w:lvl>
    <w:lvl w:ilvl="6" w:tplc="04090001" w:tentative="1">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="?"/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="5040" w:hanging="360"/>
      </w:pPr>
      <w:rPr>
        <w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/>
      </w:rPr>
    </w:lvl>
    <w:lvl w:ilvl="7" w:tplc="04090003" w:tentative="1">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="o"/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="5760" w:hanging="360"/>
      </w:pPr>
      <w:rPr>
        <w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:cs="Courier New" w:hint="default"/>
      </w:rPr>
    </w:lvl>
    <w:lvl w:ilvl="8" w:tplc="04090005" w:tentative="1">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="?"/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="6480" w:hanging="360"/>
      </w:pPr>
      <w:rPr>
        <w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/>
      </w:rPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1">
    <w:abstractNumId w:val="1"/>
  </w:num>
  <w:num w:numId="2">
    <w:abstractNumId w:val="0"/>
  </w:num>
</w:numbering>

};

declare function ooxml:settings(
) as element(w:settings)
{
<w:settings xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:sl="http://schemas.openxmlformats.org/schemaLibrary/2006/main">
  <w:zoom w:percent="100"/>
  <w:proofState w:spelling="clean" w:grammar="clean"/>
  <w:defaultTabStop w:val="720"/>
  <w:characterSpacingControl w:val="doNotCompress"/>
  <w:compat>
    <w:useFELayout/>
  </w:compat>
  <w:rsids>
    <w:rsidRoot w:val="003D2A77"/>
    <w:rsid w:val="002A2DDE"/>
    <w:rsid w:val="00397362"/>
    <w:rsid w:val="003D2A77"/>
  </w:rsids>
  <m:mathPr>
    <m:mathFont m:val="Cambria Math"/>
    <m:brkBin m:val="before"/>
    <m:brkBinSub m:val="--"/>
    <m:smallFrac m:val="off"/>
    <m:dispDef/>
    <m:lMargin m:val="0"/>
    <m:rMargin m:val="0"/>
    <m:defJc m:val="centerGroup"/>
    <m:wrapIndent m:val="1440"/>
    <m:intLim m:val="subSup"/>
    <m:naryLim m:val="undOvr"/>
  </m:mathPr>
  <w:themeFontLang w:val="en-US"/>
  <w:clrSchemeMapping w:bg1="light1" w:t1="dark1" w:bg2="light2" w:t2="dark2" w:accent1="accent1" w:accent2="accent2" w:accent3="accent3" w:accent4="accent4" w:accent5="accent5" w:accent6="accent6" w:hyperlink="hyperlink" w:followedHyperlink="followedHyperlink"/>
  <w:shapeDefaults>
    <o:shapedefaults v:ext="edit" spidmax="3074"/>
    <o:shapelayout v:ext="edit">
      <o:idmap v:ext="edit" data="1"/>
    </o:shapelayout>
  </w:shapeDefaults>
  <w:decimalSymbol w:val="."/>
  <w:listSeparator w:val=","/>
</w:settings>

};

declare function ooxml:styles(
) as element(w:styles)
{
<w:styles xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/>
        <w:sz w:val="22"/>
        <w:szCs w:val="22"/>
        <w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="ar-SA"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr>
        <w:spacing w:after="200" w:line="276" w:lineRule="auto"/>
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
  <w:latentStyles w:defLockedState="0" w:defUIPriority="99" w:defSemiHidden="1" w:defUnhideWhenUsed="1" w:defQFormat="0" w:count="267">
    <w:lsdException w:name="Normal" w:semiHidden="0" w:uiPriority="0" w:unhideWhenUsed="0" w:qFormat="1"/>
    <w:lsdException w:name="heading 1" w:semiHidden="0" w:uiPriority="9" w:unhideWhenUsed="0" w:qFormat="1"/>
    <w:lsdException w:name="heading 2" w:uiPriority="9" w:qFormat="1"/>
    <w:lsdException w:name="heading 3" w:uiPriority="9" w:qFormat="1"/>
    <w:lsdException w:name="heading 4" w:uiPriority="9" w:qFormat="1"/>
    <w:lsdException w:name="heading 5" w:uiPriority="9" w:qFormat="1"/>
    <w:lsdException w:name="heading 6" w:uiPriority="9" w:qFormat="1"/>
    <w:lsdException w:name="heading 7" w:uiPriority="9" w:qFormat="1"/>
    <w:lsdException w:name="heading 8" w:uiPriority="9" w:qFormat="1"/>
    <w:lsdException w:name="heading 9" w:uiPriority="9" w:qFormat="1"/>
    <w:lsdException w:name="toc 1" w:uiPriority="39"/>
    <w:lsdException w:name="toc 2" w:uiPriority="39"/>
    <w:lsdException w:name="toc 3" w:uiPriority="39"/>
    <w:lsdException w:name="toc 4" w:uiPriority="39"/>
    <w:lsdException w:name="toc 5" w:uiPriority="39"/>
    <w:lsdException w:name="toc 6" w:uiPriority="39"/>
    <w:lsdException w:name="toc 7" w:uiPriority="39"/>
    <w:lsdException w:name="toc 8" w:uiPriority="39"/>
    <w:lsdException w:name="toc 9" w:uiPriority="39"/>
    <w:lsdException w:name="caption" w:uiPriority="35" w:qFormat="1"/>
    <w:lsdException w:name="Title" w:semiHidden="0" w:uiPriority="10" w:unhideWhenUsed="0" w:qFormat="1"/>
    <w:lsdException w:name="Default Paragraph Font" w:uiPriority="1"/>
    <w:lsdException w:name="Subtitle" w:semiHidden="0" w:uiPriority="11" w:unhideWhenUsed="0" w:qFormat="1"/>
    <w:lsdException w:name="Strong" w:semiHidden="0" w:uiPriority="22" w:unhideWhenUsed="0" w:qFormat="1"/>
    <w:lsdException w:name="Emphasis" w:semiHidden="0" w:uiPriority="20" w:unhideWhenUsed="0" w:qFormat="1"/>
    <w:lsdException w:name="Table Grid" w:semiHidden="0" w:uiPriority="59" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Placeholder Text" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="No Spacing" w:semiHidden="0" w:uiPriority="1" w:unhideWhenUsed="0" w:qFormat="1"/>
    <w:lsdException w:name="Light Shading" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Light List" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Light Grid" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Medium Shading 1" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Medium Shading 2" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Medium List 1" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Medium List 2" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Medium Grid 1" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Medium Grid 2" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Medium Grid 3" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Dark List" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Colorful Shading" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Colorful List" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Colorful Grid" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Light Shading Accent 1" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Light List Accent 1" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Light Grid Accent 1" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Medium Shading 1 Accent 1" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Medium Shading 2 Accent 1" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Medium List 1 Accent 1" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Revision" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="List Paragraph" w:semiHidden="0" w:uiPriority="34" w:unhideWhenUsed="0" w:qFormat="1"/>
    <w:lsdException w:name="Quote" w:semiHidden="0" w:uiPriority="29" w:unhideWhenUsed="0" w:qFormat="1"/>
    <w:lsdException w:name="Intense Quote" w:semiHidden="0" w:uiPriority="30" w:unhideWhenUsed="0" w:qFormat="1"/>
    <w:lsdException w:name="Medium List 2 Accent 1" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Medium Grid 1 Accent 1" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Medium Grid 2 Accent 1" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Medium Grid 3 Accent 1" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Dark List Accent 1" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Colorful Shading Accent 1" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Colorful List Accent 1" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Colorful Grid Accent 1" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Light Shading Accent 2" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Light List Accent 2" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Light Grid Accent 2" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Medium Shading 1 Accent 2" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Medium Shading 2 Accent 2" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Medium List 1 Accent 2" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Medium List 2 Accent 2" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Medium Grid 1 Accent 2" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Medium Grid 2 Accent 2" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Medium Grid 3 Accent 2" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Dark List Accent 2" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Colorful Shading Accent 2" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Colorful List Accent 2" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Colorful Grid Accent 2" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Light Shading Accent 3" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Light List Accent 3" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Light Grid Accent 3" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Medium Shading 1 Accent 3" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Medium Shading 2 Accent 3" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Medium List 1 Accent 3" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Medium List 2 Accent 3" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Medium Grid 1 Accent 3" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Medium Grid 2 Accent 3" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Medium Grid 3 Accent 3" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Dark List Accent 3" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Colorful Shading Accent 3" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Colorful List Accent 3" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Colorful Grid Accent 3" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Light Shading Accent 4" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Light List Accent 4" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Light Grid Accent 4" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Medium Shading 1 Accent 4" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Medium Shading 2 Accent 4" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Medium List 1 Accent 4" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Medium List 2 Accent 4" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Medium Grid 1 Accent 4" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Medium Grid 2 Accent 4" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Medium Grid 3 Accent 4" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Dark List Accent 4" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Colorful Shading Accent 4" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Colorful List Accent 4" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Colorful Grid Accent 4" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Light Shading Accent 5" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Light List Accent 5" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Light Grid Accent 5" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Medium Shading 1 Accent 5" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Medium Shading 2 Accent 5" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Medium List 1 Accent 5" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Medium List 2 Accent 5" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Medium Grid 1 Accent 5" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Medium Grid 2 Accent 5" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Medium Grid 3 Accent 5" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Dark List Accent 5" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Colorful Shading Accent 5" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Colorful List Accent 5" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Colorful Grid Accent 5" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Light Shading Accent 6" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Light List Accent 6" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Light Grid Accent 6" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Medium Shading 1 Accent 6" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Medium Shading 2 Accent 6" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Medium List 1 Accent 6" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Medium List 2 Accent 6" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Medium Grid 1 Accent 6" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Medium Grid 2 Accent 6" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Medium Grid 3 Accent 6" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Dark List Accent 6" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Colorful Shading Accent 6" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Colorful List Accent 6" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Colorful Grid Accent 6" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/>
    <w:lsdException w:name="Subtle Emphasis" w:semiHidden="0" w:uiPriority="19" w:unhideWhenUsed="0" w:qFormat="1"/>
    <w:lsdException w:name="Intense Emphasis" w:semiHidden="0" w:uiPriority="21" w:unhideWhenUsed="0" w:qFormat="1"/>
    <w:lsdException w:name="Subtle Reference" w:semiHidden="0" w:uiPriority="31" w:unhideWhenUsed="0" w:qFormat="1"/>
    <w:lsdException w:name="Intense Reference" w:semiHidden="0" w:uiPriority="32" w:unhideWhenUsed="0" w:qFormat="1"/>
    <w:lsdException w:name="Book Title" w:semiHidden="0" w:uiPriority="33" w:unhideWhenUsed="0" w:qFormat="1"/>
    <w:lsdException w:name="Bibliography" w:uiPriority="37"/>
    <w:lsdException w:name="TOC Heading" w:uiPriority="39" w:qFormat="1"/>
  </w:latentStyles>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
    <w:rsid w:val="00397362"/>
  </w:style>
  <w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont">
    <w:name w:val="Default Paragraph Font"/>
    <w:uiPriority w:val="1"/>
    <w:semiHidden/>
    <w:unhideWhenUsed/>
  </w:style>
  <w:style w:type="table" w:default="1" w:styleId="TableNormal">
    <w:name w:val="Normal Table"/>
    <w:uiPriority w:val="99"/>
    <w:semiHidden/>
    <w:unhideWhenUsed/>
    <w:qFormat/>
    <w:tblPr>
      <w:tblInd w:w="0" w:type="dxa"/>
      <w:tblCellMar>
        <w:top w:w="0" w:type="dxa"/>
        <w:left w:w="108" w:type="dxa"/>
        <w:bottom w:w="0" w:type="dxa"/>
        <w:right w:w="108" w:type="dxa"/>
      </w:tblCellMar>
    </w:tblPr>
  </w:style>
  <w:style w:type="numbering" w:default="1" w:styleId="NoList">
    <w:name w:val="No List"/>
    <w:uiPriority w:val="99"/>
    <w:semiHidden/>
    <w:unhideWhenUsed/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="ListParagraph">
    <w:name w:val="List Paragraph"/>
    <w:basedOn w:val="Normal"/>
    <w:uiPriority w:val="34"/>
    <w:qFormat/>
    <w:rsid w:val="003D2A77"/>
    <w:pPr>
      <w:ind w:left="720"/>
      <w:contextualSpacing/>
    </w:pPr>
  </w:style>
</w:styles>

};

declare function ooxml:theme(
) as element(a:theme)
{
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1>
        <a:sysClr val="windowText" lastClr="000000"/>
      </a:dk1>
      <a:lt1>
        <a:sysClr val="window" lastClr="FFFFFF"/>
      </a:lt1>
      <a:dk2>
        <a:srgbClr val="1F497D"/>
      </a:dk2>
      <a:lt2>
        <a:srgbClr val="EEECE1"/>
      </a:lt2>
      <a:accent1>
        <a:srgbClr val="4F81BD"/>
      </a:accent1>
      <a:accent2>
        <a:srgbClr val="C0504D"/>
      </a:accent2>
      <a:accent3>
        <a:srgbClr val="9BBB59"/>
      </a:accent3>
      <a:accent4>
        <a:srgbClr val="8064A2"/>
      </a:accent4>
      <a:accent5>
        <a:srgbClr val="4BACC6"/>
      </a:accent5>
      <a:accent6>
        <a:srgbClr val="F79646"/>
      </a:accent6>
      <a:hlink>
        <a:srgbClr val="0000FF"/>
      </a:hlink>
      <a:folHlink>
        <a:srgbClr val="800080"/>
      </a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont>
        <a:latin typeface="Cambria"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
        <a:font script="Jpan" typeface="MS ????"/>
        <a:font script="Hang" typeface="?? ??"/>
        <a:font script="Hans" typeface="??"/>
        <a:font script="Hant" typeface="????"/>
        <a:font script="Arab" typeface="Times New Roman"/>
        <a:font script="Hebr" typeface="Times New Roman"/>
        <a:font script="Thai" typeface="Angsana New"/>
        <a:font script="Ethi" typeface="Nyala"/>
        <a:font script="Beng" typeface="Vrinda"/>
        <a:font script="Gujr" typeface="Shruti"/>
        <a:font script="Khmr" typeface="MoolBoran"/>
        <a:font script="Knda" typeface="Tunga"/>
        <a:font script="Guru" typeface="Raavi"/>
        <a:font script="Cans" typeface="Euphemia"/>
        <a:font script="Cher" typeface="Plantagenet Cherokee"/>
        <a:font script="Yiii" typeface="Microsoft Yi Baiti"/>
        <a:font script="Tibt" typeface="Microsoft Himalaya"/>
        <a:font script="Thaa" typeface="MV Boli"/>
        <a:font script="Deva" typeface="Mangal"/>
        <a:font script="Telu" typeface="Gautami"/>
        <a:font script="Taml" typeface="Latha"/>
        <a:font script="Syrc" typeface="Estrangelo Edessa"/>
        <a:font script="Orya" typeface="Kalinga"/>
        <a:font script="Mlym" typeface="Kartika"/>
        <a:font script="Laoo" typeface="DokChampa"/>
        <a:font script="Sinh" typeface="Iskoola Pota"/>
        <a:font script="Mong" typeface="Mongolian Baiti"/>
        <a:font script="Viet" typeface="Times New Roman"/>
        <a:font script="Uigh" typeface="Microsoft Uighur"/>
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface="Calibri"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
        <a:font script="Jpan" typeface="MS ??"/>
        <a:font script="Hang" typeface="?? ??"/>
        <a:font script="Hans" typeface="??"/>
        <a:font script="Hant" typeface="????"/>
        <a:font script="Arab" typeface="Arial"/>
        <a:font script="Hebr" typeface="Arial"/>
        <a:font script="Thai" typeface="Cordia New"/>
        <a:font script="Ethi" typeface="Nyala"/>
        <a:font script="Beng" typeface="Vrinda"/>
        <a:font script="Gujr" typeface="Shruti"/>
        <a:font script="Khmr" typeface="DaunPenh"/>
        <a:font script="Knda" typeface="Tunga"/>
        <a:font script="Guru" typeface="Raavi"/>
        <a:font script="Cans" typeface="Euphemia"/>
        <a:font script="Cher" typeface="Plantagenet Cherokee"/>
        <a:font script="Yiii" typeface="Microsoft Yi Baiti"/>
        <a:font script="Tibt" typeface="Microsoft Himalaya"/>
        <a:font script="Thaa" typeface="MV Boli"/>
        <a:font script="Deva" typeface="Mangal"/>
        <a:font script="Telu" typeface="Gautami"/>
        <a:font script="Taml" typeface="Latha"/>
        <a:font script="Syrc" typeface="Estrangelo Edessa"/>
        <a:font script="Orya" typeface="Kalinga"/>
        <a:font script="Mlym" typeface="Kartika"/>
        <a:font script="Laoo" typeface="DokChampa"/>
        <a:font script="Sinh" typeface="Iskoola Pota"/>
        <a:font script="Mong" typeface="Mongolian Baiti"/>
        <a:font script="Viet" typeface="Arial"/>
        <a:font script="Uigh" typeface="Microsoft Uighur"/>
      </a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office">
      <a:fillStyleLst>
        <a:solidFill>
          <a:schemeClr val="phClr"/>
        </a:solidFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0">
              <a:schemeClr val="phClr">
                <a:tint val="50000"/>
                <a:satMod val="300000"/>
              </a:schemeClr>
            </a:gs>
            <a:gs pos="35000">
              <a:schemeClr val="phClr">
                <a:tint val="37000"/>
                <a:satMod val="300000"/>
              </a:schemeClr>
            </a:gs>
            <a:gs pos="100000">
              <a:schemeClr val="phClr">
                <a:tint val="15000"/>
                <a:satMod val="350000"/>
              </a:schemeClr>
            </a:gs>
          </a:gsLst>
          <a:lin ang="16200000" scaled="1"/>
        </a:gradFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0">
              <a:schemeClr val="phClr">
                <a:shade val="51000"/>
                <a:satMod val="130000"/>
              </a:schemeClr>
            </a:gs>
            <a:gs pos="80000">
              <a:schemeClr val="phClr">
                <a:shade val="93000"/>
                <a:satMod val="130000"/>
              </a:schemeClr>
            </a:gs>
            <a:gs pos="100000">
              <a:schemeClr val="phClr">
                <a:shade val="94000"/>
                <a:satMod val="135000"/>
              </a:schemeClr>
            </a:gs>
          </a:gsLst>
          <a:lin ang="16200000" scaled="0"/>
        </a:gradFill>
      </a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
          <a:solidFill>
            <a:schemeClr val="phClr">
              <a:shade val="95000"/>
              <a:satMod val="105000"/>
            </a:schemeClr>
          </a:solidFill>
          <a:prstDash val="solid"/>
        </a:ln>
        <a:ln w="25400" cap="flat" cmpd="sng" algn="ctr">
          <a:solidFill>
            <a:schemeClr val="phClr"/>
          </a:solidFill>
          <a:prstDash val="solid"/>
        </a:ln>
        <a:ln w="38100" cap="flat" cmpd="sng" algn="ctr">
          <a:solidFill>
            <a:schemeClr val="phClr"/>
          </a:solidFill>
          <a:prstDash val="solid"/>
        </a:ln>
      </a:lnStyleLst>
      <a:effectStyleLst>
        <a:effectStyle>
          <a:effectLst>
            <a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0">
              <a:srgbClr val="000000">
                <a:alpha val="38000"/>
              </a:srgbClr>
            </a:outerShdw>
          </a:effectLst>
        </a:effectStyle>
        <a:effectStyle>
          <a:effectLst>
            <a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0">
              <a:srgbClr val="000000">
                <a:alpha val="35000"/>
              </a:srgbClr>
            </a:outerShdw>
          </a:effectLst>
        </a:effectStyle>
        <a:effectStyle>
          <a:effectLst>
            <a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0">
              <a:srgbClr val="000000">
                <a:alpha val="35000"/>
              </a:srgbClr>
            </a:outerShdw>
          </a:effectLst>
          <a:scene3d>
            <a:camera prst="orthographicFront">
              <a:rot lat="0" lon="0" rev="0"/>
            </a:camera>
            <a:lightRig rig="threePt" dir="t">
              <a:rot lat="0" lon="0" rev="1200000"/>
            </a:lightRig>
          </a:scene3d>
          <a:sp3d>
            <a:bevelT w="63500" h="25400"/>
          </a:sp3d>
        </a:effectStyle>
      </a:effectStyleLst>
      <a:bgFillStyleLst>
        <a:solidFill>
          <a:schemeClr val="phClr"/>
        </a:solidFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0">
              <a:schemeClr val="phClr">
                <a:tint val="40000"/>
                <a:satMod val="350000"/>
              </a:schemeClr>
            </a:gs>
            <a:gs pos="40000">
              <a:schemeClr val="phClr">
                <a:tint val="45000"/>
                <a:shade val="99000"/>
                <a:satMod val="350000"/>
              </a:schemeClr>
            </a:gs>
            <a:gs pos="100000">
              <a:schemeClr val="phClr">
                <a:shade val="20000"/>
                <a:satMod val="255000"/>
              </a:schemeClr>
            </a:gs>
          </a:gsLst>
          <a:path path="circle">
            <a:fillToRect l="50000" t="-80000" r="50000" b="180000"/>
          </a:path>
        </a:gradFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0">
              <a:schemeClr val="phClr">
                <a:tint val="80000"/>
                <a:satMod val="300000"/>
              </a:schemeClr>
            </a:gs>
            <a:gs pos="100000">
              <a:schemeClr val="phClr">
                <a:shade val="30000"/>
                <a:satMod val="200000"/>
              </a:schemeClr>
            </a:gs>
          </a:gsLst>
          <a:path path="circle">
            <a:fillToRect l="50000" t="50000" r="50000" b="50000"/>
          </a:path>
        </a:gradFill>
      </a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
  <a:objectDefaults/>
  <a:extraClrSchemeLst/>
</a:theme>

};

declare function ooxml:document-rels(
) as element(pr:Relationships)
{
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
  <Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>
</Relationships>
};

declare function ooxml:text(
  $text as xs:string?
) as element (w:t)
{
    <w:t xml:space="preserve">{$text}</w:t>
};

declare function ooxml:run(
  $text as element(w:t)*
) as element(w:r)
{ 
    ooxml:run($text,())
};

declare function ooxml:run( (: should $text be a + ? :)
  $text as element(w:t)*,
  $rProps as element(w:rPr)?
) as element(w:r)
{
    <w:r>
      {$rProps, $text}
    </w:r>
};

declare function ooxml:paragraph(
  $runs as element(w:r)*
) as element(w:p)
{
    ooxml:paragraph($runs,())
};

declare function ooxml:paragraph(
  $runs as element(w:r)*,
  $pProps as element(w:pPr)?
) as element(w:p)
{
    <w:p>
     {$pProps}
     {
      $runs
     }
    </w:p>
};

(: if using schema, can check for type/substitution group for what can be 
   child of $body.  need to check .xsd for type.  Or, 
   list types, check for block level children, if not acceptable, throw error
   for now doing nothing, leave function as is.   
:)
declare function ooxml:body( 
  $block-content as element()* 
) as element(w:body)
{
    <w:body>{$block-content}</w:body>
};

declare function ooxml:document(
) as element(w:document)
{
    ooxml:document(ooxml:body(()))
};

declare function ooxml:document(
  $body as element(w:body)
) as element(w:document)
{
    <w:document>
     {$body}
    </w:document>
};


(: pass in 1 for bulleted list, pass in 2 for numbering 
   subsequent paragraphs continue style (2 paras in row with num-id 2, 
   will be numbered 1. , 2. ,etc.  
   should we pass text instead of id? use "bulleted" or "numbered", 
   if only 2 options, make boolean. 
:)
declare function ooxml:list-paragraph-property( 
  $num-id as xs:string
) as element(w:pPr)
{
    <w:pPr>
       <w:pStyle w:val="ListParagraph"/>
       <w:numPr>
         <w:ilvl w:val="0"/>
         <w:numId w:val={$num-id}/>
       </w:numPr>
     </w:pPr>
};

declare function ooxml:create-simple-docx(
  $document as element(w:document)
) as binary()
{
    let $content-types := ooxml:content-types(())
    let $rels :=  ooxml:package-rels()
    let $package := ooxml:docx-package($content-types, $rels, $document)
    return $package
};

(: for now, leaving with two  ooxml:docx-package functions for 2 document
   types; simple/default.  This is good for now as docx pkg requirement are
   clear and signatures/function bodies help others to understand the formats. 
   In future may like a separate function that generates package based on 
   element node-name 
:)
declare function ooxml:docx-package(
  $content-types as element(types:Types),
  $rels          as element(pr:Relationships),
  $document      as element(w:document)
) as binary()
{
    let $manifest := <parts xmlns="xdmp:zip">
			<part>[Content_Types].xml</part>
		        <part>_rels/.rels</part>
			<part>word/document.xml</part>
		     </parts>
    let $parts := ($content-types, $rels, $document) 
    return
         xdmp:zip-create($manifest, $parts)
};

declare function ooxml:docx-package( 
  $content-types as element(types:Types),
  $rels          as element(pr:Relationships),
  $document      as element(w:document),
  $document-rels as element(pr:Relationships),
  $numbering     as element(w:numbering),
  $styles        as element(w:styles),
  $settings      as element(w:settings),
  $theme         as element(a:theme),
  $fontTable     as element(w:fonts)
) as binary()
{


    let $manifest := <parts xmlns="xdmp:zip">
			<part>[Content_Types].xml</part>
		        <part>_rels/.rels</part>
			<part>word/document.xml</part>
			<part>word/_rels/document.xml.rels</part>
			<part>word/numbering.xml</part>
			<part>word/styles.xml</part>
			<part>word/settings.xml</part>
			<part>word/theme/theme1.xml</part>
			<part>word/fontTable.xml</part>
 		     </parts>
    let $parts := ($content-types, $rels, $document, $document-rels, $numbering, $styles, $settings, $theme, $fontTable) 
    return
         xdmp:zip-create($manifest, $parts)
};


(: BEGIN replace document.xml within Flat OPC Package XML :)
declare function ooxml:passthru-pkg-doc(
  $x as node(), 
  $document-xml as element(w:document)
) as node()*
{
    for $i in $x/node() return ooxml:dispatch-doc-replace($i, $document-xml)
};

declare function ooxml:dispatch-doc-replace(
  $x as node(), 
  $document-xml as element(w:document)
) as node()?
{
    typeswitch($x)
     (: case text() return $x :) (: move document-node after element :)
     case element(w:document) return ($document-xml) 
     case element() return  element{fn:node-name($x)} {$x/@* ,passthru-pkg-doc($x, $document-xml)}
     case document-node() return document {$x/@*,ooxml:passthru-pkg-doc($x, $document-xml)}
     default return $x
};

declare function ooxml:replace-package-document(
  $pkg-xml as element(pkg:package), 
  $document-xml as element(w:document)
) as element(pkg:package)
{
   ooxml:dispatch-doc-replace($pkg-xml, $document-xml)
};
(: END replace document.xml within Flat OPC Package XML :)

declare function ooxml:docx-manifest(
  $directory as xs:string, 
  $uris      as xs:string*
) as element(zip:parts)
{
    <parts xmlns="xdmp:zip"> 
    {
      for $i in $uris
      let $dir := fn:substring-after($i,$directory)
      let $part :=  <part>{$dir}</part>
      return $part
    }
    </parts>
};
