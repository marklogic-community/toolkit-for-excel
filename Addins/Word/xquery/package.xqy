xquery version "0.9-ml"
(: Copyright 2002-2008 Mark Logic Corporation.  All Rights Reserved. :)
(: package.xqy: A library for Office OpenXML Developer Package Support      :)
module "http://marklogic.com/openxml"

declare namespace ooxml = "http://marklogic.com/openxml"
declare namespace w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
declare namespace v="urn:schemas-microsoft-com:vml"
declare namespace zip="xdmp:zip"
declare namespace pkg="http://schemas.microsoft.com/office/2006/xmlPackage"
declare namespace cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
declare namespace dc="http://purl.org/dc/elements/1.1/"
declare namespace dcterms ="http://purl.org/dc/terms/"
declare namespace rels = "http://schemas.openxmlformats.org/package/2006/relationships"

default function namespace = "http://www.w3.org/2003/05/xpath-functions"

define variable $ooxml:openxml-format-support-version { "@MAJOR_VERSION.@MINOR_VERSION@PATCH_VERSION" }

(: BEGIN USED BY PIPELINES :)
define function ooxml:get-ooxml-library-version() as xs:string
{
    $ooxml:openxml-format-support-version
}

define function ooxml:list-length-error()
{
   fn:error("ListLengthsNotEqual: ","The lengths of the lists that are dependant on each other differ.")
}

define function ooxml:validate-list-length-equal($list1 as xs:string* , $list2 as node()*) as xs:boolean
{
  fn:count($list1) eq fn:count($list2)
}


define function ooxml:validate-list-length-equal-strings($list1 as xs:string* , $list2 as xs:string*) as xs:boolean
{
  fn:count($list1) eq fn:count($list2)
}


define function ooxml:package-uris($package as node()) as xs:string*
{
   let $manifest := xdmp:zip-manifest($package)
   let $part-names := for $part-name in $manifest/zip:part return $part-name
   return $part-names
}

define function ooxml:package-parts($package as node()) as node()*
{
   let $manifest := xdmp:zip-manifest($package)
   let $parts :=
     for $part-name in $manifest/zip:part
     let $options := if (fn:ends-with($part-name, ".rels")) then
                             <options xmlns="xdmp:zip-get">
                               <format>xml</format>
                             </options>
                     else
                             <options xmlns="xdmp:zip-get"/>
     let $part := xdmp:zip-get($package, $part-name, $options)
     return $part
   return $parts
}

define function ooxml:validate-list-length-equal-2($list1 as xs:string* , $list2 as node()*) as xs:string
{
  if(fn:count($list1) eq fn:count($list2)) then "true" else ooxml:list-length-error()
}



define function ooxml:validate-directory($directory as xs:string) as xs:string
{
  let $directory-name := if(fn:ends-with($directory,"/")) then $directory else fn:concat($directory,"/")
  return $directory-name
}

define function ooxml:package-parts-insert(
  $directory as xs:string?,
  $uris as xs:string*, 
  $package-parts as node()*
) as empty() 
{
 
   let $return := if(ooxml:validate-list-length-equal($uris,$package-parts)) then 
         for $uri at $d in $uris
         let $finalname := if(fn:empty($directory)) then $uri else fn:concat(ooxml:validate-directory($directory),$uri)
         (: ADDED THIS :)
         let $cleanuri := fn:replace($finalname,"\\","/")
         return (xdmp:document-insert($cleanuri,$package-parts[$d] )) 
   else ooxml:list-length-error()
   return $return
}


define function ooxml:package-parts-insert(
  $directory as xs:string?,
  $uris as xs:string*, 
  $package-parts as node()*,
  $permissions as element(sec:permission)*
) as empty() 
{ 
   let $return := if(ooxml:validate-list-length-equal($uris,$package-parts)) then 
        for $uri at $d in $uris
        let $finalname := if(fn:empty($directory)) then $uri else fn:concat(ooxml:validate-directory($directory),$uri)
        let $cleanuri := fn:replace($finalname,"\\","/")
        return xdmp:document-insert($cleanuri,$package-parts[$d], $permissions)
   else ooxml:list-length-error()
   return $return
}

define function ooxml:package-parts-insert(
  $directory as xs:string?,
  $uris as xs:string*, 
  $package-parts as node()*,
  $permissions as element(sec:permission)*,
  $collections as xs:string*
) as empty() 
{ 
   let $return := if(ooxml:validate-list-length-equal($uris,$package-parts)) then 
       for $uri at $d in $uris
       let $finalname := if(fn:empty($directory)) then $uri else fn:concat(ooxml:validate-directory($directory),$uri)
       let $cleanuri := fn:replace($finalname,"\\","/")
       return xdmp:document-insert($cleanuri,$package-parts[$d], $permissions, $collections)
   else ooxml:list-length-error()
   return $return
}

define function ooxml:package-parts-insert(
  $directory as xs:string?,
  $uris as xs:string*, 
  $package-parts as node()*,
  $permissions as element(sec:permission)*,
  $collections as xs:string*,
  $quality as xs:int
) as empty() 
{ 
   let $return := if(ooxml:validate-list-length-equal($uris,$package-parts)) then 
       for $uri at $d in $uris
       let $finalname := if(fn:empty($directory)) then $uri else fn:concat(ooxml:validate-directory($directory),$uri)
       let $cleanuri := fn:replace($finalname,"\\","/")
       return xdmp:document-insert($cleanuri,$package-parts[$d], $permissions, $collections, $quality)
   else ooxml:list-length-error()
   return $return
}


define function ooxml:package-parts-insert(
  $directory as xs:string?,
  $uris as xs:string*, 
  $package-parts as node()*,
  $permissions as element(sec:permission)*,
  $collections as xs:string*,
  $quality as xs:int,
  $forest-ids as xs:unsignedLong*
) as empty() 
{
   let $return := if(ooxml:validate-list-length-equal($uris,$package-parts)) then 
      for $uri at $d in $uris
      let $finalname := if(fn:empty($directory)) then $uri else fn:concat(ooxml:validate-directory($directory),$uri)
      let $cleanuri := fn:replace($finalname,"\\","/")
      return xdmp:document-insert($cleanuri,$package-parts[$d], $permissions, $collections, $quality, $forest-ids)
   else ooxml:list-length-error()
   return $return
}
(: END USED BY PIPELINES :)
