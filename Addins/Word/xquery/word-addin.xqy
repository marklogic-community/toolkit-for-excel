xquery version "0.9-ml"
(: Copyright 2002-2008 Mark Logic Corporation.  All Rights Reserved. :)
(: addin-lib.xqy: A library for Office OpenXML Developer Support      :)
module "addin-lib"
import module namespace scx = "custom-xml-lib" at "/MarkLogic/openxml/custom-xml.xqy"

declare namespace mla = "addin-lib"
declare namespace w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
declare namespace v="urn:schemas-microsoft-com:vml"
declare namespace zip="xdmp:zip"
declare namespace pkg="http://schemas.microsoft.com/office/2006/xmlPackage"
declare namespace cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
declare namespace dc="http://purl.org/dc/elements/1.1/"
declare namespace dcterms ="http://purl.org/dc/terms/"
declare namespace rels = "http://schemas.openxmlformats.org/package/2006/relationships"

default function namespace = "http://www.w3.org/2003/05/xpath-functions"

define variable $mla:openxml-format-support-version { "@MAJOR_VERSION.@MINOR_VERSION@PATCH_VERSION" }

(: BEGIN USED IN PIPELINES (also in openxml.xqy) ========================================================== :)
define function mla:get-mla-library-version() as xs:string
{
    $mla:openxml-format-support-version
}

define function mla:list-length-error()
{
   fn:error("ListLengthsNotEqual: ","The lengths of the lists that are dependant on each other differ.")
}

define function mla:validate-list-length-equal($list1 as xs:string* , $list2 as node()*) as xs:boolean
{
  fn:count($list1) eq fn:count($list2)
}


define function mla:validate-list-length-equal-strings($list1 as xs:string* , $list2 as xs:string*) as xs:boolean
{
  fn:count($list1) eq fn:count($list2)
}

define function mla:validate-list-length-equal-2($list1 as xs:string* , $list2 as node()*) as xs:string
{
  if(fn:count($list1) eq fn:count($list2)) then "true" else mla:list-length-error()
}
(: END USED BY PIPELINES (also in openxml.xqy) =============================================================:)


(: BEGIN w:customXml HIGHLIGHT --------------------------- :)
define function mla:custom-xml-highlight($nodes as node()*, $highlight-term as cts:query, $tag-name as xs:string,  $attributes as xs:string*, $values as xs:string*) as  node()*
{
   let $return := if(mla:validate-list-length-equal-strings($attributes,$values)) then 
      scx:custom-xml-highlight($nodes,$highlight-term,$tag-name, $attributes, $values)
   else mla:list-length-error()
   return $return
}

define function mla:custom-xml-highlight($nodes as node()*, $highlight-term as cts:query, $tag-name as xs:string) as  node()*
{
      scx:custom-xml-highlight($nodes,$highlight-term,$tag-name)
}
(: END w:customXml HIGHLIGHT ---------------------------  :)

(: BEGIN REMOVE w:p PROPERTIES -------------------------- :)
define function mla:passthru-para($x as node()*) as node()*
{
   for $i in $x/node() return mla:dispatch-paragraph-to-clean($i)
}

define function mla:dispatch-paragraph-to-clean($x as node()*) as node()*
{
  
    typeswitch ($x)
     case element(w:pPr) return ()
     case element(w:rPr) return () 
     case text() return $x
     default return (  
       element{fn:node-name($x)} {$x/@*,mla:passthru-para($x)} 
     )
}

define function mla:remove-paragraph-properties($paragraph as node()*) as node()*
{
    mla:dispatch-paragraph-to-clean($paragraph)
}

(: END REMOVE w:p PROPERTIES ---------------------------- :)

(: BEGIN ADDIN-SEARCH HELPERS ---------------------------  :)
define function mla:get-customxml-parent($doc as node()*) as node()*
{

   if($doc/parent::w:sdtContent) then mla:get-customxml-parent($doc/../..) 
   else if($doc/parent::w:customXml) then mla:get-customxml-parent($doc/..)
   else $doc
 
}

define function mla:insure-insertable-paragraph ($x as node()*) as node()*
{
   if(fn:exists($x//w:p) or ($x/self::w:p)) then $x else <w:p>{$x}</w:p>
}

define function mla:paragraph-search($query as cts:query) as node()*
{
    let $doc := cts:search(//w:p ,$query)
    return $doc
}


define function mla:paragraph-search($query as cts:query, $begin as xs:integer, $end as xs:integer) as node()*
{
    let $doc := cts:search(//w:p ,$query)[$begin to $end]
    return $doc
}

define function mla:customxml-search($query as cts:query, $element as xs:string) as node()*
{
  let $cust := cts:search(//w:customXml[@w:element=$element], $query)
  return $cust
}

define function mla:customxml-search($query as cts:query, $element as xs:string,$begin as xs:integer, $end as xs:integer) as node()*
{
  let $cust := cts:search(//w:customXml[@w:element=$element], $query)[$begin to $end]
  return $cust
}

define function mla:sdt-search($query as cts:query, $element as xs:string) as node()*
{
   let $sdt   := cts:search(//w:sdt[w:sdtPr/w:tag/@w:val=$element],($query))
   return $sdt
}

define function mla:sdt-search($query as cts:query, $element as xs:string,$begin as xs:integer, $end as xs:integer) as node()*
{
   let $sdt   := cts:search(//w:sdt[w:sdtPr/w:tag/@w:val=$element],($query))[$begin to $end]
   return $sdt
}

define function mla:custom-search-all($query as cts:query, $begin as xs:integer, $end as xs:integer) as node()*
{
    let $sdt := cts:search( //(w:sdt | w:customXml | w:p ), ($query))[$begin to $end]
    return $sdt
}

(: END ADDIN-SEARCH HELPERS ---------------------------  :)

(: BEGIN REPLACE PARAGRAPH FORMATTING ---------------------------  :)

define function mla:update-paragraph-format-passthru($x as node()*, $wpProps as element(w:pPr)?, $wrProps as element(w:rPr)?) as node()*
{
       for $i in $x/node() return mla:update-paragraph-format-dispatch($i, $wpProps, $wrProps)
}

define function mla:update-paragraph-format-dispatch($wp as node()*, $wpProps as element(w:pPr)?, $wrProps as element(w:rPr)?) as node()*
{
       typeswitch ($wp)
       case text() return $wp
       case document-node() return document {$wp/@*,mla:update-paragraph-format-passthru($wp, $wpProps, $wrProps)}
       case element(w:p) return mla:add-paragraph-properties($wp, $wpProps, $wrProps)
       case element(w:r) return mla:add-run-style-properties($wp, $wrProps) 
       case element() return  element{fn:node-name($wp)} {$wp/@*,mla:update-paragraph-format-passthru($wp, $wpProps, $wrProps)}
       default return $wp
}

define function mla:add-run-style-properties($wr as node(),$runprops as element(w:rPr)? ) as node()
{
       element w:r{ $runprops, $wr/node() }
}
(: removed style : was: add-paragraph-style-properties :)
define function mla:add-paragraph-properties($wp as node()*, $paraprops as element(w:pPr)?, $runprops as element(w:rPr)?) as node()*
{
        element w:p{ $paraprops, mla:update-paragraph-format-passthru($wp, $paraprops, $runprops) }
}

  (: called by mla-addin-update-paragraph-format.xqy  :)  
define function mla:update-paragraph-format($wp as node()*, $wpProps as element(w:pPr)?, $wrProps as element(w:rPr)?) as node()*
{
    let $newpara := mla:update-paragraph-format-dispatch($wp, $wpProps, $wrProps) 
    return $newpara
}


(: END REPLACE PARAGRAPH FORMATTING ---------------------------  :)

define function mla:add-style($styles as node(), $newstyle as node()*) as node()
{
  element w:styles { $styles/@*,
                     $styles/node(),
                     $newstyle
  }
}

(: END USED BY ADDIN  ========================================================== :)
(: BEGIN ADDED FOR LATEST ADDIN ====================================================== :)

define function mla:transfer-paragraph-properties($styleparagraph as element(w:p), $contentparagraph as element(w:p)) as element(w:p)
{
    let $tgtpara := $contentparagraph
    let $srcpara := $styleparagraph


    let $t1 := mla:remove-paragraph-properties($tgtpara)
    let $o1 := $srcpara

    (:multiple paragraphs could be sent, get style from last one :)
    let $pcount := fn:count($o1)
    let $paraprops := $o1[$pcount]/w:pPr[1]

    (:same for runs, get last w:rPr from last w:r in w:p :)
    let $rcount := fn:count($o1[$pcount]/w:r)
    let $rprcount := fn:count($o1[$pcount]/w:r[$rcount]/w:rPr)
    let $runprops:= $o1[$pcount]/w:r[$rcount]/w:rPr[$rprcount]

    let $newpara := mla:update-paragraph-format($t1, $paraprops, $runprops) 

    return $newpara


}

define function mla:update-styles-xml($styleIds as xs:string, $sourcedocuri as xs:string, $activedocstyles as node()) as node()*
{
   let $currentstyles := $activedocstyles
   let $styleId := fn:tokenize($styleIds,",")
   let $styleupd := for $s in $styleId
                    let $x :=  if(fn:empty($currentstyles//w:style[@w:styleId=$styleId])) then  
                                         let $stylesource := doc(fn:concat($sourcedocuri,"/word/styles.xml"))
                                         let $newstyle := $stylesource/w:styles/w:style[@w:styleId=$s]
                                         return $newstyle
                                         else ()
                     return $x
                                          
   return  if(fn:empty($styleupd)) then () else mla:add-style($currentstyles, $styleupd)


}
(: END ADDED FOR LATEST ADDIN ======================================================== :)



(: BEGIN ADD NUMBERING.XML TO PKG R&D ========================================================== :)

define function mla:updDocRels($currels as node(), $newmax as xs:string) as node()
{
element pkg:part { $currels/@*, 
        element pkg:xmlData {
               element Relationships{
                                      $currels/pkg:xmlData/rels:Relationships/@*,
                                      $currels/pkg:xmlData/rels:Relationships/node(),
                                      <Relationship Id={$newmax} Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>


                                    }
                            }         
                 }

}


define function mla:add-numberxml-to-docx($docxml as node(), $numbering as node()*) as node()
{ (:
 element pkg:package { $docxml/@*,
                       $docxml/node(),
                       element {      }  //HERE
                     }
   :)




let $docx := $docxml
(: REPLACE BOTTOM TWO WITH NUMBERS.XML IN FUNCTION :)
let $numbers := $numbering
let $newpkgpart :=
              <pkg:part pkg:name="/word/numbering.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml">
                 <pkg:xmlData>{$numbers}</pkg:xmlData>
              </pkg:part>

let $currentpkgrels := $docx/pkg:package/pkg:part[@pkg:name="/_rels/.rels"]
let $currentdocrels := $docx/pkg:package/pkg:part[@pkg:name="/word/_rels/document.xml.rels"](:/pkg:xmlData/rels:Relationships:)
let $currelids :=  $currentdocrels/pkg:xmlData/rels:Relationships/rels:Relationship/@Id

let $curmax := fn:max(for $i in $currelids 
                      return fn:substring-after(fn:string($i),"rId"))

let $newmax := fn:concat("rId",($curmax cast as xs:integer)+1)

let $updatedpkg :=
  element pkg:package {
                         $docx/@*,
                         $currentpkgrels,
                         mla:updDocRels($currentdocrels,$newmax),
                         $docx/pkg:package/pkg:part[fn:not(@pkg:name="/word/_rels/document.xml.rels") and fn:not(@pkg:name="/_rels/.rels")],
                         $newpkgpart
                         
                         
                         
                      }

let $newpkg := $updatedpkg
return  $newpkg  

}

(: END ADD NUMBERING.XML TO PKG R&D ========================================================== :)



(: BEGIN CREATION OF <pkg:package> ========================================================== :)

define function mla:get-part-content-type($uri as xs:string) as xs:string?
{
   if(fn:ends-with($uri,".rels"))
   then "application/vnd.openxmlformats-package.relationships+xml"
   else if(fn:ends-with($uri,"document.xml"))
   then
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
   else if(fn:matches($uri, "theme\d+\.xml"))
   then "application/vnd.openxmlformats-officedocument.theme+xml"
   else if(fn:ends-with($uri,"word/numbering.xml"))
   then "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"
   (: else if(fn:ends-with($uri,"word/settings.xml")):)
   else if(fn:ends-with($uri,"settings.xml"))
   then "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"
   (:else if(fn:ends-with($uri,"word/styles.xml")):)
   else if(fn:ends-with($uri,"styles.xml"))
   then "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"
   (: else if(fn:ends-with($uri,"word/webSettings.xml")) :)
   else if(fn:ends-with($uri,"webSettings.xml"))
   then "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml"
   (: else if(fn:ends-with($uri,"word/fontTable.xml")) :)
   else if(fn:ends-with($uri,"fontTable.xml"))
   then "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"
   else if(fn:ends-with($uri,"word/footnotes.xml"))
   then "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"
   else if(fn:matches($uri, "header\d+\.xml"))
   then 
        "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
   else if(fn:matches($uri, "footer\d+\.xml"))
   then "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"
   else if(fn:ends-with($uri,"word/endnotes.xml"))
   then
        "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"
   else if(fn:ends-with($uri,"docProps/core.xml"))
   then
        "application/vnd.openxmlformats-package.core-properties+xml"
   else if(fn:ends-with($uri,"docProps/app.xml"))
   then
       "application/vnd.openxmlformats-officedocument.extended-properties+xml"
   else if(fn:ends-with($uri,"docProps/custom.xml")) then
       "application/vnd.openxmlformats-officedocument.custom-properties+xml"
   else if(fn:ends-with($uri,"jpeg")) then
        "image/jpeg"
   else if(fn:ends-with($uri,"wmf")) then
        "image/x-wmf"
   else if(fn:matches($uri,"customXML/itemProps\d+\.xml")) then
        "application/vnd.openxmlformats-officedocument.customXmlProperties+xml"
   else if(fn:matches($uri,"customXML/item\d+\.xml")) then
        "application/xml"
   else
       ()
    
}

define function mla:get-part-attributes($uri as xs:string) as node()*
{
  (:not sure if this is needed, for serverside generated docx, path comes through as \path\name , if path has /mixed\separators\for\path/example, it chokes when opening in word :)
  let $cleanuri := fn:replace($uri,"\\","/")
  let $name := attribute pkg:name{$cleanuri}
  let $contenttype := attribute pkg:contentType{mla:get-part-content-type($cleanuri)}
  let $padding := if(fn:ends-with($cleanuri,".rels")) then
                     if(fn:starts-with($cleanuri,"/_rels")) then
                      attribute pkg:padding{ "512" }
                     else
                      attribute pkg:padding{ "256" }
                  else
                     ()
  let $compression := if(fn:ends-with($cleanuri,"jpeg")) then 
                         attribute pkg:compression { "store" } 
                      else ()
  
  return ($name, $contenttype, $padding, $compression)
}

define function mla:get-package-part($directory as xs:string, $uri as xs:string) as node()?
{
  let $fulluri := $uri
  let $docuri := fn:concat("/",fn:substring-after($fulluri,$directory))
  let $data := doc($fulluri)

  let $part := if(fn:empty($data) or fn:ends-with($fulluri,"[Content_Types].xml")) then () 
               else if(fn:ends-with($fulluri,".jpeg") or fn:ends-with($fulluri,".wmf")) then 
                  element pkg:part { mla:get-part-attributes($docuri), element pkg:binaryData {   xs:base64Binary(xs:hexBinary($data))    }   }
               else
                  element pkg:part { mla:get-part-attributes($docuri), element pkg:xmlData { $data }}
  return  $part (: <T>{$fulluri}</T>   :) 
}

define function mla:make-package($directory as xs:string, $uris as xs:string*) as node()
{
  let $package := element pkg:package { 
                            for $uri in $uris
                            let $part := mla:get-package-part($directory,$uri)
                            return $part }
                           
  return $package
}


define function mla:package-uris-from-directory($docuri as xs:string) as xs:string*
{

  cts:uris("","document",cts:directory-query($docuri,"infinity"))

}

(: USAGE

let $directory:="/CNN3.docx/"
let $uris := mlos:package-uris-from-directory($directory)
  (:remove dirs from list :)
let $validuris := for $uri in $uris
                  let $u := if(fn:ends-with($uri,"/")) then () else fn:concat("/",$uri)
                  return $u
return     mla:make-package($directory, $validuris) 


:)

(: END CREATION OF <pkg:package> ========================================================== :)









(: NOT QUITE READY FOR PRIME TIME   :)
define function mla:get-run-cxml($text as xs:string, $tag as xs:string, $attrs as xs:string*, $vals as xs:string*) as element(w:customXml)
{
(:Question : better to break up into get-para-cxml, get-run-cxml, or on function with both pass the option :)
(: element constructor? or is this ok? :)

  <w:customXml w:element={$tag}>
    { if(fn:count($attrs) gt 1 )
      then
       <w:customXmlPr>
         {
           for $attr at $d in $attrs 
           return
            <w:attr w:name ={$attr}  w:val={$vals[$d]} />
         }
       </w:customXmlPr>
       else ()
    } 
    <w:r>
     <w:t>{$text}</w:t>
    </w:r>
  </w:customXml>

}

define function mla:get-para-cxml($text as xs:string, $tag as xs:string, $attrs as xs:string*, $vals as xs:string*) as element(w:customXml)
{
  <w:customXml w:element={$tag}>
      { if(fn:count($attrs) gt 1 )
      then
       <w:customXmlPr>
         {
           for $attr at $d in $attrs 
           return
            <w:attr w:name ={$attr}  w:val={$vals[$d]} />
         }
       </w:customXmlPr>
       else ()
    }
     <w:p>
     <w:r>
       <w:t>{$text}</w:t>
     </w:r>
   </w:p>
  </w:customXml>

}
