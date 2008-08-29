xquery version "0.9-ml"
(: Copyright 2002-2008 Mark Logic Corporation.  All Rights Reserved. :)
(: word-processing-ml.xqy: library for Word support, includes:
   merging runs
   custom-xml-highlight
   styles support for paragraphs and styles.xml
:)

module "http://marklogic.com/openxml"

declare namespace w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
declare namespace v="urn:schemas-microsoft-com:vml"
declare namespace ve="http://schemas.openxmlformats.org/markup-compatibility/2006"
declare namespace o="urn:schemas-microsoft-com:office:office"
declare namespace r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
declare namespace m="http://schemas.openxmlformats.org/officeDocument/2006/math"
declare namespace wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
declare namespace w10="urn:schemas-microsoft-com:office:word" 
declare namespace wne="http://schemas.microsoft.com/office/word/2006/wordml"

declare namespace ooxml = "http://marklogic.com/openxml"
import module "http://marklogic.com/openxml" at "/MarkLogic/openxml/package.xqy"

define variable $ooxml:wml-format-support-version { "@MAJOR_VERSION.@MINOR_VERSION@PATCH_VERSION"}
 
define function ooxml:ooxml-version() as xs:string
{
    (: $ooxml:wordprocessingml-format-support-version :)
    $ooxml:wml-format-support-version
}

(: START MERGE RUNS ========================================================================== :)
(: define function update-document-xml($document as element(w:document)) as element(w:document) :)
define function update-document-xml($document as node()*) as node()*
{
  let $doc := $document (:/element() :)
  return  dispatch($doc)

}

define function passthru($x as node()*) as node()*
{
for $i in $x/node() return dispatch($i)
}

define function dispatch ($x as node()*) as node()*
{ (:checkfor optimization, had if($x//w:p) then typeswitch else $x but  then not all runs merged, booooooo! :)

     typeswitch ($x)
       case text() return $x
       case document-node() return document {$x/@*,passthru($x)}
       case element(w:p) return mergeruns($x) 
       case element() return  element{fn:name($x)} {$x/@*,passthru($x)}
       default return $x
}

define function mergeruns($p as element(w:p)) as element(w:p)
{
    let $rsidR := $p/@w:rsidR
    let $rsidRDefault := $p/@w:rsidRDefault
    let $pPrvals := if(fn:exists($p/w:pPr)) then $p/w:pPr else ()

    return  element w:p{ $p/@*, $pPrvals,
                        if(fn:count($pPrvals) gt 0 ) then map($p/w:pPr/following-sibling::*[1]) else map($p/child::*[1]) }
}

define function descend($r as node()?, $rToCheck as element(w:rPr)?) as element(w:r)*
{
    let $uri := "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    let $nodename := fn:node-name($r)
    let $ret :=

      if(fn:empty($r)) then ()
      else if($nodename eq fn:QName($uri,"r")) then 
      if(fn:deep-equal($r/w:rPr,$rToCheck))
           then
              ($r, descend($r/following-sibling::*[1], $rToCheck))
           else ()
      else ()
    return $ret
}

define function map($r as node()*) as node()*
{
    let $uri := "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    let $nodename := fn:node-name($r) 
     return

       if (fn:empty ($r)) then () 
       else if($nodename eq fn:QName($uri,"r")) then

         let $nextsib := $r/following-sibling::*[1]
         let $rToCheck := $r/w:rPr

          (: have to account for embeds; EXAMPLE: (table;w:r/w:pict/v:shape/v:textbox/w:txbxContent/w:p/w:r)) :)
          let $wpEmbed := $r//w:p
          let $newPict := if(fn:empty($wpEmbed)) then () 
                          else dispatch($r) 

          let $matches := descend($nextsib, $rToCheck)
          let $count := fn:count($matches)

          let $this := if ($count) then 
                          (element w:r{ $rToCheck, 
                           element w:t { attribute xml:space{"preserve"}, fn:string-join(($r/w:t, $matches/w:t),"") } }) 
                       else if (fn:not(fn:empty($wpEmbed))) then
                          (element w:r{ $rToCheck, $newPict} )  
                       else $r
 
          return  ($this, 
             if($count) 
             then ($r/following-sibling::*[1+$count],map($r/following-sibling::*[2+$count]))
             else ((), map($r/following-sibling::*[1]) ) 
                  ) 
       else (element{fn:name($r)}  {$r/@*,map($r/child::*[1])} ,map($r/following-sibling::*[1])) 
}

define function ooxml:runs-merge($nodes as node()*) as node()*
{
   ooxml:update-document-xml($nodes)
}

(: END MERGE RUNS ============================================================================ :)

define function ooxml:create-paragraph($para as xs:string) as element(w:p)
{
  element w:p{ element w:r { element w:t {$para}}}
}

(: BEGIN REMOVE w:p PROPERTIES =============================================================== :)
define function ooxml:passthru-para($x as node()) as node()
{
   for $i in $x/node() return ooxml:dispatch-paragraph-to-clean($i)
}

define function ooxml:dispatch-paragraph-to-clean($x as node()) as node()
{

      typeswitch($x)
       case text() return $x
       case document-node() return document {$x/@*,ooxml:passthru-para($x)}
       case element(w:pPr) return ()
       case element(w:rPr) return () 
       case element() return  element{fn:name($x)} {$x/@*,passthru-para($x)}
       default return $x

}

define function ooxml:remove-paragraph-styles($paragraph as element()) as element()
{
    ooxml:dispatch-paragraph-to-clean($paragraph)
}

(: END REMOVE w:p PROPERTIES ================================================================= :)
define function ooxml:get-paragraph-styles($paragraph as element(w:p)*) as element(w:pPr)*
{
   $paragraph//w:pPr
}

define function ooxml:get-run-styles($paragraph as element(w:p)*) as element(w:rPr)*
{
   $paragraph//w:rPr
}

define function ooxml:get-paragraph-style-id($pstyle as element (w:pPr)) as xs:string?
{
   let $styles := $pstyle//w:pStyle/@w:val
   return $styles 
}

define function ooxml:get-run-style-id($rstyle as element (w:rPr)) as xs:string?
{
   let $styles := $rstyle//w:rStyle/@w:val
   return $styles 
}

define function ooxml:get-style-definition($styleid as xs:string, $styles as element(w:styles) ) as element(w:style)?
{
   for $id in $styleid 
   return $styles//w:style[@w:styleId=$id]
}

define function ooxml:replace-style-definition($newstyle as element(w:style), $styles as element(w:styles)) as element(w:styles)
{
                 element w:styles { $styles/@*,
                     $styles/* except $styles//w:style[@w:styleId=$newstyle/@w:styleId],
                     $newstyle }
}



(: BEGIN SET PARAGRAPH STYLES ================================================================  :)



define function ooxml:set-paragraph-styles-passthru($x as node()*, $props as element()?, $type as xs:string) as node()*
{
       for $i in $x/node() return ooxml:set-paragraph-styles-dispatch($i, $props, $type)
}

define function ooxml:set-paragraph-styles-dispatch($wp as node()*, $props as element()?, $type as xs:string ) as node()*
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
          

}

define function ooxml:add-run-style-properties($wr as node(),$runprops as element(w:rPr)? ) as node()*
{
       element w:r{ $wr/@*, $runprops, $wr/* except $wr/w:rPr }
}

define function ooxml:add-paragraph-properties($wp as node()*, $paraprops as element(w:pPr)?, $type as xs:string) as node()*
{
        element w:p{ $wp/@*, $paraprops, element w:r { $wp/w:r/@*, ooxml:set-paragraph-styles-passthru($wp/* except $wp/w:pPr, $paraprops, $type) }}
}

define function ooxml:replace-paragraph-styles($block as element(), $wpProps as element(w:pPr)?) as element()
{
     ooxml:set-paragraph-styles-dispatch($block,$wpProps,"wp")
}

define function ooxml:replace-run-styles($block as element(), $wrProps as element(w:rPr)?) as element()
{
     ooxml:set-paragraph-styles-dispatch($block,$wrProps,"wr")
}


(:
define function ooxml:set-paragraph-styles($wp as node()*, $wpProps as element(w:pPr)?, $wrProps as element(w:rPr)?) as node()*
{ 
    let $clean := ooxml:remove-paragraph-styles($wp)
    let $newpara := ooxml:set-paragraph-styles-dispatch($clean, $wpProps, $wrProps) 
    return $newpara
}

:)

(: END SET PARAGRAPH STYLES ==================================================================== :)
(: TBD :)
define function ooxml:transfer-paragraph-styles($styleparagraph as element(w:p), $contentparagraph as element(w:p)) as element(w:p)
{

}

define function ooxml:custom-xml($content as element(), $tag as xs:string) as element(w:customXml)
{
  (:check element passed in that it can be child of w:customXml :)

  element w:customXml{attribute w:element{$tag}, $content}
}

(: BEGIN SET CUSTOM XML TAG ==================================================================== :)
define function ooxml:set-custom-xml-passthru($x as node()*, $oldtag as xs:string, $newtag as xs:string) as node()*
{
       for $i in $x/node() return ooxml:set-custom-xml-dispatch($i, $oldtag, $newtag)
}

define function ooxml:set-custom-xml-dispatch($block as node()*, $oldtag as xs:string, $newtag as xs:string) as node()*
{
       typeswitch ($block)
       case text() return $block
       case document-node() return document {$block/@*,ooxml:set-custom-xml-passthru($block, $oldtag, $newtag)}
       case element(w:customXml) return ooxml:set-custom-element-value($block, $oldtag, $newtag) 
       case element() return  element{fn:node-name($block)} {$block/@*,ooxml:set-custom-xml-passthru($block, $oldtag, $newtag)}
       default return $block
}

define function ooxml:set-custom-element-value($block as node()*, $oldtag as xs:string, $newtag as xs:string) as node()*
{
   let $value := $block/@w:element
   let $cxml := if($value eq $oldtag) then
                      element w:customXml {attribute w:element{$newtag}, $block/@* except $block/@w:element, ooxml:set-custom-xml-passthru($block, $oldtag, $newtag)}
                   else
                      element{fn:node-name($block)} {$block/@*,ooxml:set-custom-xml-passthru($block, $oldtag, $newtag)} 
   return $cxml
}

define function ooxml:replace-custom-xml-element($content as element(), $oldtag as xs:string, $newtag as xs:string) as element()
{ 
    let $newblock := ooxml:set-custom-xml-dispatch($content, $oldtag, $newtag) 
    return $newblock
}
(: END SET CUSTOM XML TAG ====================================================================== :)

define function ooxml:get-custom-xml-ancestor($doc as element()) as element()?
{

   if($doc/parent::w:sdtContent) then ooxml:get-custom-xml-ancestor($doc/../..) 
   else if($doc/parent::w:customXml) then ooxml:get-custom-xml-ancestor($doc/..)
   else $doc
 
}

(: BEGIN SIMPLE SEARCH ================================================================================ :)

define function ooxml:paragraph-search($query as cts:query) as node()*
{
    let $doc := cts:search(//w:p ,$query)
    return $doc
}

define function ooxml:paragraph-search($query as cts:query, $begin as xs:integer, $end as xs:integer) as node()*
{
    let $doc := cts:search(//w:p ,$query)[$begin to $end]
    return $doc
}

define function ooxml:custom-search-all($query as cts:query, $begin as xs:integer, $end as xs:integer) as node()*
{
    let $sdt := cts:search( //(w:sdt | w:customXml | w:p ), ($query))[$begin to $end]
    return $sdt
}

(: END SIMPLE SEARCH ================================================================================== :)

(: BEGIN w:customXml HIGHLIGHT ================================================================= :)

define function ooxml:passthru-chlt($x as node()*) as node()*
{
  for $i in $x/node() return ooxml:dispatch-chlt($i)
}

define function ooxml:map($props as node()*, $x as node()*) as node()*
{

  for $child in $x return
   typeswitch ($child)
    case text() return ooxml:makerun($child, $props)
    case element(w:customXml) return element{fn:name($child)} {$child/@*, $child/w:customXmlPr, <w:r>{$props,$child/w:r/child::*}</w:r>}
    case element() return element{fn:name($x)} {$x/@*,ooxml:passthru-chlt($x)}
    default return $x
}


define function ooxml:dispatch-chlt($x as node()*) as node()*
{
   typeswitch ($x)
    case document-node() return ooxml:passthru-chlt($x)
    case text() return $x
    case element(w:r) return (if(fn:exists($x//child::*//w:p)) then ooxml:passthru-chlt($x) 
                              else ooxml:map(<w:rPr>{$x/w:rPr/node()}</w:rPr>, $x/w:t/node()))
    case element() return  element{fn:name($x)} {$x/@*,ooxml:passthru-chlt($x)} 
    default return $x
}

define function ooxml:makerun($x as text(), $runProps as element(w:rPr)) as element(w:r)
{ 
    <w:r>{$runProps}<w:t xml:space="preserve">{$x}</w:t></w:r>
}

define function ooxml:custom-xml-highlight-exec($orig as node()*, $query as cts:query, $tagname as xs:string, $attrs as xs:string*, $vals as xs:string*) as node()*
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
}

define function ooxml:custom-xml-highlight-exec($orig as node()*, $query as cts:query, $tagname as xs:string) as node()*
{    let $tmpdoc := <temp>{$orig}</temp>
     let $highlightedbody := cts:highlight($tmpdoc, $query, 
                               <w:customXml w:element="{$tagname}">
                                    <w:r><w:t>{$cts:text}</w:t></w:r>
                               </w:customXml>)
     let $newdocument := ooxml:dispatch-chlt($highlightedbody)
     return $newdocument/*
}


define function ooxml:custom-xml-highlight($nodes as node()*, $highlight-term as cts:query, $tag-name as xs:string,  $attributes as xs:string*, $values as xs:string*) as  node()*
{
   let $return := if(ooxml:validate-list-length-equal-strings($attributes,$values)) then 
      ooxml:custom-xml-highlight-exec($nodes,$highlight-term,$tag-name, $attributes, $values)
   else ooxml:list-length-error()
   return $return
}

define function ooxml:custom-xml-highlight($nodes as node()*, $highlight-term as cts:query, $tag-name as xs:string) as  node()*
{
      ooxml:custom-xml-highlight-exec($nodes,$highlight-term,$tag-name)
}

(: END w:customXml HIGHLIGHT =================================================================== :)

