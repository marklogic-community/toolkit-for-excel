xquery version "1.0-ml";
declare namespace html = "http://www.w3.org/1999/xhtml";
declare namespace test = "http://test";
declare namespace p= "http://schemas.openxmlformats.org/presentationml/2006/main";
declare namespace r="http://schemas.openxmlformats.org/officeDocument/2006/relationships";
declare namespace rel="http://schemas.openxmlformats.org/package/2006/relationships";
declare namespace w="http://schemas.openxmlformats.org/wordprocessingml/2006/main";

import module namespace excel= "http://marklogic.com/openxml/excel" 
       at "/MarkLogic/openxml/spreadsheet-ml-support.xqy";

import module namespace ooxml="http://marklogic.com/openxml" 
	    at "/MarkLogic/openxml/word-processing-ml-support.xqy";

import module namespace ppt=  "http://marklogic.com/openxml/powerpoint" at "/MarkLogic/openxml/presentation-ml-support.xqy";

<tests>
<test1>
{
(: will open in compatability mode in 2010 :)
 let $doc := ooxml:document(ooxml:body(ooxml:create-paragraph("Hello, World!")))
 return xdmp:save("C:\simpleDocxTest-1.docx",ooxml:create-simple-docx($doc))
}
</test1>
<test2>
{
(: will open in compatability mode in 2010 :)
let $text2 := ooxml:text("FOOBAR2")
let $run := ooxml:run($text2)
let $para := ooxml:paragraph($run)
let $body := ooxml:body($para)
let $document := ooxml:document($body)
return xdmp:save("C:\simpleDocxTest-2.docx",ooxml:create-simple-docx($document))
}
</test2>
<test3>
{
(: will open in compatability mode in 2010 :)
let $text2 := ooxml:text("FOOBAR3")
let $run := ooxml:run($text2)
let $para := ooxml:paragraph($run)
let $body := ooxml:body($para)
let $document := ooxml:document($body)
let $content-types := ooxml:simple-content-types()
let $rels := ooxml:package-rels()
return xdmp:save("C:\simpleDocxTest-3.docx",ooxml:docx-package($content-types,$rels,$document))
}
</test3>
<test4>
{
(:open in correct mode for each 2007 and 2010 :)
let $content-types:= ooxml:default-content-types()
let $rels := ooxml:package-rels()
let $para := ooxml:create-paragraph("Hello, World!")
let $doc := ooxml:document(ooxml:body($para))

let $doc-rels := ooxml:document-rels()
let $numbering := ooxml:numbering()
let $styles := ooxml:styles()
let $settings := ooxml:settings()
let $theme := ooxml:theme()
let $font-table := ooxml:font-table() 

return xdmp:save("C:\simpleDocxTest-4.docx",ooxml:docx-package($content-types, $rels, $doc, $doc-rels, $numbering, $styles, $settings, $theme, $font-table))

}
</test4>
<test5>
{
(:assumes sampleManuscript.docx has been saved to db and extracted using Office Open XML Extract pipeline :)
(: ooxml:package has been updated to add 2010 namespaces to certain .docx components
   so 2007 documents may open properly in 2010; 2007 docs will open in 2010 in compatability mode :)
let $directory:="/sampleManuscript_docx_parts/" 
let $package := ooxml:get-directory-package($directory)
return xdmp:save("C:\testOutput-1.xml",$package)
}
</test5>
<test6>
{
let $directory := "/sampleManuscript_docx_parts/"
let $uris := ooxml:directory-uris($directory)
let $parts := fn:doc(($uris))/node()
return xdmp:save("C:\testOutput-2.xml",ooxml:package($parts))
}
</test6>
<test7>
{
let $text2 := ooxml:text("Four scored and seven years ago, our forefathers put forth on this great nation...")
let $run := ooxml:run($text2)
let $para := ooxml:paragraph($run)
let $para2 := ooxml:create-paragraph("this is a test")
let $body := ooxml:body(($para,$para2))
let $document := ooxml:document($body)
let $directory := "/sampleManuscript2010_docx_parts/"
let $uris := ooxml:directory-uris($directory)
let $parts := fn:doc(($uris))/node()
let $package := ooxml:package($parts)
return xdmp:save("C:\testOutput10-3.xml",ooxml:replace-package-document($package,$document))

}
</test7>
<test8>
{
(:testing package with images :)
let $directory:="/TractorDrivers2007_docx_parts/" 
let $package := ooxml:get-directory-package($directory)
return xdmp:save("C:\testOutput-4.xml",$package)
}
</test8>
<test9>
{
let $directory:="/TractorDrivers2010_docx_parts/" 
let $img:= fn:doc("/TractorDrivers2007_docx_parts/word/media/image1.jpeg")/node()
return xdmp:save("C:\testOutput-5.xml", <base64>
                  {ooxml:base64-opc-format(ooxml:binary-to-base64-string($img))}
                 </base64>)
}
</test9>
<test10>
{
let $directory:="/TractorDrivers2010_docx_parts/" 
let $img:= fn:doc("/TractorDrivers2007_docx_parts/word/media/image1.jpeg")/node()
return xdmp:save("C:\testOutput-6.jpeg",
                   ooxml:base64-string-to-binary(
                      ooxml:base64-opc-format(
                        ooxml:binary-to-base64-string($img)
                      )
                   )
                 )
}
</test10>  
<test11>
{
(: customXml can be used to enrich extracted documents and applied to both 2007 and 2010, though it will not be exposed in the Word interface and won't be retained on subsequent save. :)
let $text2 := ooxml:text("CUSTOMXML TEST")
let $run := ooxml:run($text2)
let $para := ooxml:paragraph($run)
let $tag := "MyTag"
let $body := ooxml:body(ooxml:custom-xml($para,$tag))
let $document := ooxml:document($body)
return xdmp:save("C:\simpleDocxTest-4.docx",ooxml:create-simple-docx($document))
}
</test11>
<test12>
{
(: customXml can be used to enrich extracted documents and applied to both 2007 and 2010, though it will not be exposed in the Word interface and won't be retained on subsequent save. :)
let $wp := <w:p><w:r><w:t>The MarkLogic government summit was in Washington DC this year.</w:t></w:r></w:p>
let $body := ooxml:body(ooxml:custom-xml-entity-highlight($wp))
let $document := ooxml:document($body)
return xdmp:save("C:\simpleDocxTest-5.docx",ooxml:create-simple-docx($document))
}
</test12>
<test13>
{
let $wp := <w:p><w:r><w:t>THIS IS A TEST!</w:t></w:r></w:p>
let $wordquery := cts:word-query("TEST")
let $body := ooxml:body(ooxml:custom-xml-highlight($wp,$wordquery,"MyTag"))
let $document := ooxml:document($body)
return xdmp:save("C:\simpleDocxTest-6.docx",ooxml:create-simple-docx($document))
}
</test13>
<test14>
{
xdmp:save("C:\testOutput-7.xml",
              <directory-uris>{
                ooxml:directory-uris("/TractorDrivers2007_docx_parts/")
                              }</directory-uris>)
}
</test14>
<test15>
{
xdmp:save("C:\testOutput-8.xml",
              <mimetype>{
              ooxml:get-mimetype("/sampleManuscript2010.docx")
                        }</mimetype>)

}
</test15>
<test16>
{
let $styleids := "DefaultParagraphFont"
let $styles := ooxml:styles() 
let $definition := ooxml:get-style-definition($styleids, $styles)
return

xdmp:save("C:\testOutput-9.xml",
              <style-definition>{
                            $definition
                        }</style-definition>)

}
</test16>
<test17>
{
let $block:=
  <w:customXml w:element="ABCD">
    <w:p>
       <w:r><w:t>This is my paragraph.</w:t></w:r>
    </w:p>
  </w:customXml>
let $origtag := "ABCD"
let $newtag := "EFGH"
return xdmp:save("C:\testOutput-10.xml",ooxml:replace-custom-xml-element($block,$origtag,$newtag))
}
</test17>
<test18>
{
let $wp1 := 
<w:p>
     <w:pPr><w:u/></w:pPr>
     <w:r><w:t>THIS IS A TEST!</w:t></w:r>
</w:p>
let $pPr := <w:pPr><w:b/></w:pPr>
let $body:=ooxml:body(ooxml:replace-paragraph-styles($wp1,$pPr))
let $document := ooxml:document($body)
return xdmp:save("C:\simpleDocxTest-7.docx",ooxml:create-simple-docx($document))

}
</test18>
<test19>
{
let $wp1 := 
<w:p>
     <w:r><w:t>THIS IS A TEST!</w:t></w:r>
</w:p>
let $rPr := <w:rPr><w:b/></w:rPr>
let $body:=ooxml:body(ooxml:replace-run-styles($wp1,$rPr))
let $document := ooxml:document($body)
return xdmp:save("C:\simpleDocxTest-8-a.docx",ooxml:create-simple-docx($document)) 
}
</test19>
<test20>
{
let $wp1 := 
<w:p>
     <w:r><w:rPr><w:u/></w:rPr>
     <w:t>THIS IS A TEST!</w:t></w:r>
</w:p>
let $rPr := <w:rPr><w:b/></w:rPr>
let $body:=ooxml:body(ooxml:replace-run-styles($wp1,$rPr))
let $document := ooxml:document($body)
return xdmp:save("C:\simpleDocxTest-8-b.docx",ooxml:create-simple-docx($document)) 
}
</test20>
</tests>
