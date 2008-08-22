xquery version "0.9-ml"
(: Copyright 2002-2008 Mark Logic Corporation.  All Rights Reserved. :)
(: custom-xml-lib.xqy: A library for Office OpenXML Developer Support, 
:: specifically addressing use of <w:customXml> and <w:sdt>          :)
module "custom-xml-lib"
declare namespace scx = "custom-xml-lib"

declare namespace w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
declare namespace v="urn:schemas-microsoft-com:vml"
declare namespace ve="http://schemas.openxmlformats.org/markup-compatibility/2006"
declare namespace o="urn:schemas-microsoft-com:office:office"
declare namespace r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
declare namespace m="http://schemas.openxmlformats.org/officeDocument/2006/math"
declare namespace wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
declare namespace w10="urn:schemas-microsoft-com:office:word" 
declare namespace wne="http://schemas.microsoft.com/office/word/2006/wordml"

define function passthru($x as node()*) as node()*
{
  for $i in $x/node() return dispatch($i)
}

define function map($props as node()*, $x as node()*) as node()*
{

  for $child in $x return
   typeswitch ($child)
    case text() return makerun($child, $props)
    case element(w:customXml) return element{fn:name($child)} {$child/@*, $child/w:customXmlPr, <w:r>{$props,$child/w:r/child::*}</w:r>}
    case element() return element{fn:name($x)} {$x/@*,passthru($x)}
    default return $x
}


define function dispatch($x as node()*) as node()*
{
   typeswitch ($x)
    case document-node() return passthru($x)
    case text() return $x
    case element(w:r) return (if(fn:exists($x//child::*//w:p)) then passthru($x) 
                              else map(<w:rPr>{$x/w:rPr/node()}</w:rPr>, $x/w:t/node()))
    case element() return  element{fn:name($x)} {$x/@*,passthru($x)} 
    default return $x
}

define function makerun($x as text(), $runProps as element(w:rPr)) as element(w:r)
{ 
    <w:r>{$runProps}<w:t xml:space="preserve">{$x}</w:t></w:r>
}

(: define function set-custom-xml-document($origdoc as element(w:document)?, $texttomarkup as xs:string, $tagname as xs:string) as element(w:document)? :)
(: define function set-custom-xml-document($orig as node()*, $texttomarkup as xs:string, $tagname as xs:string) as node()* :)
define function custom-xml-highlight($orig as node()*, $query as cts:query, $tagname as xs:string, $attrs as xs:string*, $vals as xs:string*) as node()*
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
     let $newdocument := dispatch($highlightedbody)
     (: return $highlightedbody :)
     return $newdocument/*
}

define function custom-xml-highlight($orig as node()*, $query as cts:query, $tagname as xs:string) as node()*
{    let $tmpdoc := <temp>{$orig}</temp>
     let $highlightedbody := cts:highlight($tmpdoc, $query, 
                               <w:customXml w:element="{$tagname}">
                                    <w:r><w:t>{$cts:text}</w:t></w:r>
                               </w:customXml>)
     let $newdocument := dispatch($highlightedbody)
     return $newdocument/*
}


