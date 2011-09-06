xquery version "1.0-ml";
(:
Copyright 2008-2011 MarkLogic Corporation

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

declare namespace search="http://marklogic.com/openxml/search";
declare namespace w    ="http://schemas.openxmlformats.org/wordprocessingml/2006/main";
declare namespace q    ="http://marklogic.com/beta/searchbox";
declare namespace xlink="http://www.w3.org/1999/xlink";
declare namespace pkg="http://schemas.microsoft.com/office/2006/xmlPackage";
declare namespace ins="http://marklogic.com/openxml/insert";
declare namespace dc="http://purl.org/dc/elements/1.1/";
(:office 2010:)
(: declare namespace mc = "http://schemas.openxmlformats.org/markup-compatibility/2006"; :)
import module namespace ooxml= "http://marklogic.com/openxml" at 
                               "/MarkLogic/openxml/word-processing-ml-support.xqy";

declare function ins:passthru-pkg-doc(
  $pkg as node(), 
  $body-xml as element(w:body)
) as node()*
{
    for $i in $pkg/node() return ins:dispatch-body-replace($i, $body-xml)
};

declare function ins:dispatch-body-replace(
  $pkg as node(), 
  $new-body as element(w:body)
) as node()*
{
    typeswitch($pkg)
     case text() return $pkg
     case document-node() return document {$pkg/@*, ins:passthru-pkg-doc($pkg, $new-body)}
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
                                           {$pkg/@*,ins:passthru-pkg-doc($pkg, $new-body)}
                                       </w:document>

     case element(w:styles) return <w:styles xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
                                             xmlns:o="urn:schemas-microsoft-com:office:office" 
                                             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" 
                                             xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" 
                                             xmlns:v="urn:schemas-microsoft-com:vml" 
                                             xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" 
                                             xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" 
                                             xmlns:w10="urn:schemas-microsoft-com:office:word"
                                             xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" 
                                             xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" 
                                             xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" 
                                             xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" 
                                             xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
                                     >
                                           {$pkg/@*,ins:passthru-pkg-doc($pkg, $new-body)}
                                   </w:styles>
     case element(w:fonts) return  <w:fonts xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" 
                                            xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
                                   >
                                           {$pkg/@*,ins:passthru-pkg-doc($pkg, $new-body)}
                                   </w:fonts>
     case element(w:webSettings) return <w:webSettings xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" 
                                                       xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
                                        >
                                           {$pkg/@*,ins:passthru-pkg-doc($pkg, $new-body)}
                                        </w:webSettings>
     case element(w:settings) return <w:settings xmlns:o="urn:schemas-microsoft-com:office:office"
                                                 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                                                 xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
                                                 xmlns:v="urn:schemas-microsoft-com:vml"
                                                 xmlns:w10="urn:schemas-microsoft-com:office:word"
                                                 xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
                                                 xmlns:sl="http://schemas.openxmlformats.org/schemaLibrary/2006/main"
                                      >
                                           {$pkg/@*,ins:passthru-pkg-doc($pkg, $new-body)}
                                      </w:settings>
     case element(w:glossaryDocument) return <w:glossaryDocument xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" 
                                                                 xmlns:o="urn:schemas-microsoft-com:office:office" 
                                                                 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" 
                                                                 xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" 
                                                                 xmlns:v="urn:schemas-microsoft-com:vml" 
                                                                 xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" 
                                                                 xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" 
                                                                 xmlns:w10="urn:schemas-microsoft-com:office:word" 
                                                                 xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" 
                                                                 xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" 
                                                                 xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" 
                                                                 xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" 
                                                                 xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
                                              >
                                           {$pkg/@*,ins:passthru-pkg-doc($pkg, $new-body)}
                                              </w:glossaryDocument>
     case element(w:body) return ($new-body) 
     case element() return  element{fn:node-name($pkg)} {$pkg/@* ,ins:passthru-pkg-doc($pkg, $new-body)} 
     default return $pkg
};

declare function ins:passthru-sdt-doc(
  $sdt as node()
) as node()*
{
    for $i in $sdt/node() return ins:dispatch-sdtContent-replace($i)
};

declare function ins:dispatch-sdtContent-replace(
  $sdt as node()
) as node()*
{
    typeswitch($sdt)
     case text() return $sdt
     case document-node() return document {$sdt/@*, ins:passthru-sdt-doc($sdt)}
      case element(w:r) return <w:p>{$sdt}</w:p>
     case element() return  element{fn:node-name($sdt)} {$sdt/@* ,ins:passthru-sdt-doc($sdt)} 
   
     default return $sdt
};


let $uri := xdmp:get-request-field("uri")
let $path := xdmp:get-request-field("path")
let $doc := fn:doc($uri)
let $sdt := $doc/xdmp:unpath($path)
(:let $log := xdmp:log(fn:concat("===================================",fn:string(fn:node-name($sdt/w:sdtContent/node()[1])))):)
let $content := if (fn:node-name($sdt/w:sdtContent/node()[1]) eq fn:QName($ooxml:WORDPROCESSINGML,"r") or
                     fn:empty($sdt/w:sdtContent/w:p)) then
                       ins:dispatch-sdtContent-replace($sdt) 
                else 
                       $sdt
let $body := <w:body>{$content}</w:body>

(:check to see if this came from extracted .docx, or Word saved as XML :)
let $pkg:= if(fn:contains($uri,"_docx_parts")) then
                   ooxml:get-directory-package(fn:substring-before($uri,"word/document.xml"))
            else $doc
let $upd-pkg :=  ins:dispatch-body-replace($pkg, $body)   (: the pkg:package for insert :)

let $metadata-ids := ($body//w:id/@w:val) (: returned in order :)
let $metadata := for $id in $metadata-ids
                     return if(fn:not(fn:empty($pkg//dc:metadata[dc:identifier eq $id]))) then 
                                 $pkg//dc:metadata[dc:identifier eq $id]
                            else <dc:metadata/>
(:let $save := xdmp:save("C:\testSDT.xml",$upd-pkg):)
return xdmp:quote(<insertable>
                     <insertpkg>{$upd-pkg}</insertpkg>
                     <meta>{$metadata}</meta>
                  </insertable>)

