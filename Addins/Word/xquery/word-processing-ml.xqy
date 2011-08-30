xquery version "1.0-ml";
(: Copyright 2002-2011 MarkLogic Corporation.  All Rights Reserved. :)
(: package.xqy: A library for Office OpenXML Developer Package Support      :)
module namespace ooxml = "http://marklogic.com/openxml";

(:Office 2007:)
declare namespace w="http://schemas.openxmlformats.org/wordprocessingml/2006/main";
declare namespace v="urn:schemas-microsoft-com:vml";
declare namespace ve="http://schemas.openxmlformats.org/markup-compatibility/2006";
declare namespace o="urn:schemas-microsoft-com:office:office";
declare namespace r="http://schemas.openxmlformats.org/officeDocument/2006/relationships";
declare namespace m="http://schemas.openxmlformats.org/officeDocument/2006/math";
declare namespace wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
declare namespace w10="urn:schemas-microsoft-com:office:word";
declare namespace wne="http://schemas.microsoft.com/office/word/2006/wordml";

(:Office 2010:)
declare namespace a="http://schemas.openxmlformats.org/drawingml/2006/main";
declare namespace pic="http://schemas.openxmlformats.org/drawingml/2006/picture";
declare namespace wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas";
declare namespace mc="http://schemas.openxmlformats.org/markup-compatibility/2006";
declare namespace wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing";
declare namespace w14="http://schemas.microsoft.com/office/word/2010/wordml";
declare namespace wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup";
declare namespace wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk";
declare namespace wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape";


import module "http://marklogic.com/openxml" at "/MarkLogic/openxml/package.xqy";

(: START MERGE RUNS ========================================================================== :)
declare function update-document-xml($document as node()*) as node()*
{
  let $doc := $document (:/element() :)
  return  dispatch($doc)

};

declare function passthru($x as node()*) as node()*
{
for $i in $x/node() return dispatch($i)
};

declare function dispatch ($x as node()*) as node()*
{
(: adding namespaces for Office 2010.  This wasn't required for 2007, where you only had to declare the namespaces
   actually used in document.xml.  2010 requires these, even if not used in document.xml.  
   No harm in using with Office 2007 as it will discard what's not used when it consumes the XML :)
     typeswitch ($x)
       case text() return $x
       case document-node() return document {$x/@*,passthru($x)}
       case element(w:document) return <w:document xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                                                   xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
                                                   xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
                                                   xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
                                                   xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
                                                   xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
                                                   xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
                                                   xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"
                                                   xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
                                        >
                                           {$x/@*,passthru($x)}
                                       </w:document> 
       case element(w:p) return mergeruns($x) 
       case element() return  element{fn:name($x)} {$x/@*,passthru($x)}
       default return $x
};

declare function mergeruns($p as element(w:p)) as element(w:p)
{
    let $rsidR := $p/@w:rsidR
    let $rsidRDefault := $p/@w:rsidRDefault
    let $pPrvals := if(fn:exists($p/w:pPr)) then $p/w:pPr else ()

    return  element w:p{ $p/@*, $pPrvals,
                        if(fn:count($pPrvals) gt 0 ) then map($p/w:pPr/following-sibling::*[1]) else map($p/child::*[1]) }
};

declare function descend($r as node()?, $rToCheck as element(w:rPr)?) as element(w:r)*
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
};

declare function map($r as node()*) as node()*
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
             then map($r/following-sibling::*[1+$count])
             else ((), map($r/following-sibling::*[1]) ) 
                  ) 
       else (element{fn:name($r)}  {$r/@*,map($r/child::*[1])} ,map($r/following-sibling::*[1])) 
};

declare function ooxml:runs-merge($nodes as node()*) as node()*
{
   ooxml:update-document-xml($nodes)
};

(: END MERGE RUNS ============================================================================ :)

