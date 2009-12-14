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
let $text2 := ooxml:text("FOOBAR2")
let $run := ooxml:run($text2)
let $para := ooxml:paragraph($run)
let $body := ooxml:body($para)
let $document := ooxml:document($body)
return xdmp:save("C:\tmp\xqTestOutput\HelloWorld.docx",ooxml:create-simple-docx($document))
}
</test1>
<test2>
{
let $text2 := ooxml:text("FOOBAR2")
let $run := ooxml:run($text2)
let $para := ooxml:paragraph($run)
let $body := ooxml:body($para)
let $document := ooxml:document($body)
let $content-types := ooxml:simple-content-types()
let $rels := ooxml:package-rels()
return xdmp:save("C:\tmp\xqTestOutput\HelloWorld-2.docx",ooxml:docx-package($content-types,$rels,$document))
}
</test2>
<test3>
{
let $text2 := ooxml:text("FOOBAR2")
let $run := ooxml:run($text2)
let $para := ooxml:paragraph($run)
let $body := ooxml:body($para)
let $document := ooxml:document($body)
let $content-types := ooxml:default-content-types()
let $rels := ooxml:package-rels()
let $document-rels := ooxml:document-rels()
let $numbering := ooxml:numbering()
let $styles := ooxml:styles()
let $settings := ooxml:settings()
let $theme := ooxml:theme()
let $font-table := ooxml:font-table()
return xdmp:save("C:\tmp\xqTestOutput\HelloWorld-3.docx",ooxml:docx-package($content-types,$rels,$document, $document-rels, $numbering, $styles, $settings, $theme, $font-table))
}
</test3>
<test4>
{
let $directory:="/sampleManuscript2_docx_parts/" 
let $package := ooxml:get-directory-package($directory)
return xdmp:save("C:\tmp\xqTestOutput\HelloWorld-4.xml",$package)
}
</test4>
<test5>
{
let $directory := "/sampleManuscript2_docx_parts/"
let $uris := ooxml:directory-uris($directory)
let $validuris := ooxml:package-files-only($uris)
let $parts := fn:doc(($uris))/node()
return xdmp:save("C:\tmp\xqTestOutput\HelloWorld-5.xml",ooxml:package($parts))
}
</test5>
<test6>
{
let $text2 := ooxml:text("FOOBAR2")
let $run := ooxml:run($text2)
let $para := ooxml:paragraph($run)
let $para2 := ooxml:create-paragraph("this is a test")
let $body := ooxml:body(($para,$para2))
let $document := ooxml:document($body)
let $directory := "/sampleManuscript2_docx_parts/"
let $uris := ooxml:directory-uris($directory)
let $validuris := ooxml:package-files-only($uris)
let $parts := fn:doc(($uris))/node()
let $package := ooxml:package($parts)
return xdmp:save("C:\tmp\xqTestOutput\HelloWorld-6.xml",ooxml:replace-package-document($package,$document))
}
</test6>
<test7>
{
let $directory:="/imagetest3_docx_parts/" 
let $package := ooxml:get-directory-package($directory)
return xdmp:save("C:\tmp\xqTestOutput\HelloWorld-7.xml",$package)
}
</test7>
<test8>
{
let $directory := "/imagetest_docx_parts/"
let $uris := ooxml:directory-uris($directory)
let $validuris := ooxml:package-files-only($uris)
let $parts := fn:doc(($uris))/node()
return xdmp:save("C:\tmp\xqTestOutput\HelloWorld-8.xml",ooxml:package($parts))
}
</test8>
</tests>
