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

module namespace config="http://marklogic.com/toolkit/excel/author/config";
declare namespace dc="http://purl.org/dc/elements/1.1/";


declare variable $config:CONFIG-PATH := "http://localhost:8030/xlAuthor/config/";
declare variable $config:USER := "oslo";
declare variable $config:PWD  := "oslo";

(: security discussion in notes :)
declare variable $config:TAGS :=         config:get-config-document("tags.xml");
declare variable $config:METADATA :=     config:get-config-document("metadata.xml");
declare variable $config:SEARCH :=       config:get-config-document("search.xml");
declare variable $config:COMPARE :=      config:get-config-document("compare.xml");

declare function config:get-config-document($type as xs:string)
{
 xdmp:document-get(fn:concat($config:CONFIG-PATH, $type),
                              <options xmlns="xdmp:document-get"
                                       xmlns:http="xdmp:http">
                                <format>xml</format>
                                <http:authentication>
	                           <http:username>{$config:USER}</http:username>
	                           <http:password>{$config:PWD}</http:password>
	                        </http:authentication>
                              </options>)
};

(:BEGIN Current-Document - TAGs Tab Display:)
declare function config:workbook-tags()
{
    let $img-inline := $config:TAGS/node()/config:workbook/config:tag
    for $t at $d in $img-inline
    (:let $func := fn:concat("workbookTagFunc",$d,"()"):)
    let $func := fn:concat("workbookTagFunc",$d,"()")
    return <li>
             <a href="#" onmouseup="blurSelected(this)" onclick="{$func}">{$t/config:name/text()}</a>
           </li>


};

declare function config:worksheet-tags()
{
    let $img-inline := $config:TAGS/node()/config:worksheet/config:tag
    for $t at $d in $img-inline
    let $func := fn:concat("worksheetTagFunc",$d,"()")
    return <li>
             <a href="#" onmouseup="blurSelected(this)" onclick="{$func}">{$t/config:name/text()}</a>
           </li>


};

declare function config:component-tags()
{
    let $img-inline := $config:TAGS/node()/config:component/config:tag
    for $t at $d in $img-inline
    let $func := fn:concat("componentTagFunc",$d,"()")
    return <li>
             <a href="#" onmouseup="blurSelected(this)" onclick="{$func}">{$t/config:name/text()}</a>
           </li>


};
(:END Current-Document - TAGs Tab Display:)

(:BEGIN Current-Document - Tags Tab - Generate Javascript Functions  :)
(: update checkWorkbookTags, 
   setWorkbookProperties after  var wbName = :)
(: relation is either workbook, worksheet, or component
   type is name of workbook, worksheet, or componenet
   identifier is tag applied 
   description is used for chart serizization and other :)
declare function config:generate-js-workbook-tag-func()
{
    let $wb-tags := $config:TAGS/node()/config:workbook/config:tag
    for $wbt at $d in $wb-tags
    let $name := $wbt/config:name/text() (:display label:)
    let $value := $wbt/config:value/text()  (:used for name:)
    return fn:concat("function workbookTagFunc",$d,"(){",
                       "var tagId = randomId();",
                       "if(!(checkWorkbookTags('",$value,"'))){;",

                       "var wbName = MLA.getActiveWorkbookName();",
                       "var stringxml = MLA.unescapeXMLCharEntities(generateTemplate(map.get('",$value,"')));",
                       "var domxml = MLA.createXMLDOM(stringxml);",
                       "var source = domxml.getElementsByTagName('dc:source')[0];",
                       "var relation = domxml.getElementsByTagName('dc:relation')[0];",
                       "var type = domxml.getElementsByTagName('dc:type')[0];",
                       "var id = domxml.getElementsByTagName('dc:identifier')[0];",

                       "if(source.hasChildNodes())",
	                   "{",
		              "source.childNodes[0].nodeValue='';",
	 	              "source.childNodes[0].nodeValue=tagId;",
	                   "}", 
	                   "else",
	                   "{",
	                      "var child = source.appendChild(domxml.createTextNode(tagId));",
	               "}",


                       "if(relation.hasChildNodes())",
	                   "{",
		              "relation.childNodes[0].nodeValue='';",
	 	              "relation.childNodes[0].nodeValue='workbook';",
	                   "}", 
	                   "else",
	                   "{",
	                      "var child = relation.appendChild(domxml.createTextNode('workbook'));",
	               "}",


	               "if(type.hasChildNodes())",
	                   "{",
		              "type.childNodes[0].nodeValue='';",
	 	              "type.childNodes[0].nodeValue=wbName;",
	                   "}", 
	                   "else",
	                   "{",
	                      "var child = type.appendChild(domxml.createTextNode(wbName));",
	               "}",

                       "if(id.hasChildNodes())",
	                       "{",
		                 "id.childNodes[0].nodeValue='';",
	 	                 "id.childNodes[0].nodeValue='",$value,"';",
	                       "}", 
	                         "else",
	                       "{",
	                          "var child = id.appendChild(domxml.createTextNode('",$value,"'));",
	               "}",

                           "MLA.addCustomXMLPart(domxml.xml);",
                           "setWorkbookProperties();",

                       (:"alert(stringxml+'FOO'+domxml.xml);:) 
                           "}",
	
                  "}") 
};

declare function  config:generate-js-worksheet-tag-func()
{
    let $sheet-tags := $config:TAGS/node()/config:worksheet/config:tag
    for $st at $d in $sheet-tags
    let $name := $st/config:name/text()
    let $value := $st/config:value/text()
    return fn:concat("function worksheetTagFunc",$d,"(){",
                      "var tagId = randomId();",
                      "if(!(checkWorksheetTags('",$value,"'))){;",

                      "var wsName = MLA.getActiveWorksheetName();",
                      "var stringxml = MLA.unescapeXMLCharEntities(generateTemplate(map.get('",$value,"')));",
                      "var domxml = MLA.createXMLDOM(stringxml);",
                      "var source = domxml.getElementsByTagName('dc:source')[0];",
                      "var relation = domxml.getElementsByTagName('dc:relation')[0];",
                      "var type = domxml.getElementsByTagName('dc:type')[0];",
                      "var id = domxml.getElementsByTagName('dc:identifier')[0];",

                      "if(source.hasChildNodes())",
	                   "{",
		              "source.childNodes[0].nodeValue='';",
	 	              "source.childNodes[0].nodeValue=tagId;",
	                   "}", 
	                   "else",
	                   "{",
	                      "var child = source.appendChild(domxml.createTextNode(tagId));",
	               "}",

                       "if(relation.hasChildNodes())",
	                   "{",
		              "relation.childNodes[0].nodeValue='';",
	 	              "relation.childNodes[0].nodeValue='worksheet';",
	                   "}", 
	                   "else",
	                   "{",
	                      "var child = relation.appendChild(domxml.createTextNode('worksheet'));",
	               "}",


	               "if(type.hasChildNodes())",
	                   "{",
		              "type.childNodes[0].nodeValue='';",
	 	              "type.childNodes[0].nodeValue=wsName;",
	                   "}", 
	                   "else",
	                   "{",
	                      "var child = type.appendChild(domxml.createTextNode(wsName));",
	               "}",

                       "if(id.hasChildNodes())",
	                       "{",
		                 "id.childNodes[0].nodeValue='';",
	 	                 "id.childNodes[0].nodeValue='",$value,"';",
	                       "}", 
	                         "else",
	                       "{",
	                          "var child = id.appendChild(domxml.createTextNode('",$value,"'));",
	               "}",

                           "MLA.addCustomXMLPart(domxml.xml);",
                           "setWorksheetProperties();",

                       (:"alert(stringxml+'FOO'+domxml.xml);:) 
                           "}",
	
                  "}") 
};
(: check component tags - is it named range or chart?
added getSelectedRangeName and getSelectedChartName
if selected chartname, then chart
if selectedRangeCoords, then named range
   apply namedrange tag by adding namedrange, if namedrange and info to customxml part
   apply charttag by adding info to customxml part, including base64 serialization of image
:)
declare function  config:generate-js-component-tag-func()
{
    let $component-tags := $config:TAGS/node()/config:component/config:tag
    for $ct at $d in $component-tags
    let $name := $ct/config:name/text()
    let $value := $ct/config:value/text()
    return fn:concat("function componentTagFunc",$d,"(){",
                       "var tagId = randomId();",
                       "var componentName = '';",
                       "var componentRelation = '';",
                       "var description = '';",
                       "if(!(checkComponentTags('",$value,"'))){",
                       (:"var idx = MLA.getSlideIndex();",:)
                          "try{", 
                               "var chartName = MLA.getSelectedChartName();",
                               "if(!(chartName==null || chartName==''))",
                                "{",
                                     (: "alert('chartName'+chartName);", :)
		                     "componentName=chartName;",
		                     "componentRelation='chart';",
                                     (: have to check custom part doesn't already exist for tagged item,
                                        if so, delete custom part, then add new part.
                                        this mimics excel's behavior.  you can only have a component with a unique name in each sheet,
                                        but you can reassigne the name to another named range, and the prior one ceases to exist. :)
                                     "var sheetName = MLA.getActiveWorksheetName();",
                                     (: "alert('sheetName '+sheetName);", :)
                                     "var newName = trim(chartName.substring(sheetName.length,chartName.length));",
                                     (: "alert('newName '+newName);", :)
                                     "var tmpPath = MLA.getTempPath()+newName+'.PNG';",
                                      (: "alert('tmpPath '+tmpPath);", :) 

 
	                             "var success = MLA.exportChartImagePNG(tmpPath);",
                                     (: "alert('success '+success);", :)
	   
	                             "description =MLA.base64EncodeImage(tmpPath);",
                                     (: "alert('chartImage: '+description);", :) 
                                     "var deleted = MLA.deleteFile(tmpPath);",
                                      (:get image serialization:)
	                        "}",
                                "else",
                                "{",
                                  "var rangeCoords = MLA.getSelectedRangeCoordinates();",   
                                    (: "alert('rangeCoords'+rangeCoords);", :) 
                                "if(!(rangeCoords==null || rangeCoords==''))",
                                "{",
                                   "var startCell='';",
                                   "var endCell='';",
		                   "var range = rangeCoords.split(':');",
                                   "startCell=range[0];",
                                    "if(range[1]==null || range[1] =='')",
                                    "{",
		                     "endCell=range[0];",
                                      (:get image serialization:)
	                            "}",
                                    "else",
                                    "{",
		                     "endCell=range[1];",
                                      (:get image serialization:)
	                            "}",
                                    "var sheetName = MLA.getActiveWorksheetName();",
                                    "var namedRange = MLA.addNamedRange(startCell, endCell,'",$value,"',sheetName);",
                                    "componentName = MLA.getSelectedRangeName();",
                                    "description = MLA.getSelectedRangeCoordinates();",
                                     
                                    "componentRelation = 'namedrange';",
	                         "}",
                                 "}", 
                                        
                       "var stringxml = MLA.unescapeXMLCharEntities(generateTemplate(map.get('",$value,"')));",
                       "var domxml = MLA.createXMLDOM(stringxml);",
                       "var source = domxml.getElementsByTagName('dc:source')[0];",
                       "var relation = domxml.getElementsByTagName('dc:relation')[0];",
                       "var type = domxml.getElementsByTagName('dc:type')[0];",
                       "var id = domxml.getElementsByTagName('dc:identifier')[0];",
                       "var desc = domxml.getElementsByTagName('dc:description')[0];",

                      "if(source.hasChildNodes())",
	                   "{",
		              "source.childNodes[0].nodeValue='';",
	 	              "source.childNodes[0].nodeValue=tagId;",
	                   "}", 
	                   "else",
	                   "{",
	                      "var child = source.appendChild(domxml.createTextNode(tagId));",
	               "}",

                       "if(relation.hasChildNodes())",
	                   "{",
		              "relation.childNodes[0].nodeValue='';",
	 	              "relation.childNodes[0].nodeValue=componentRelation;",
	                   "}", 
	                   "else",
	                   "{",
	                      "var child = relation.appendChild(domxml.createTextNode(componentRelation));",
	               "}",


	               "if(type.hasChildNodes())",
	                   "{",
		              "type.childNodes[0].nodeValue='';",
	 	              "type.childNodes[0].nodeValue=componentName;",
	                   "}", 
	                   "else",
	                   "{",
	                      "var child = type.appendChild(domxml.createTextNode(componentName));",
	               "}",

                       "if(id.hasChildNodes())",
	                       "{",
		                 "id.childNodes[0].nodeValue='';",
	 	                 "id.childNodes[0].nodeValue='",$value,"';",
	                       "}", 
	                         "else",
	                       "{",
	                          "var child = id.appendChild(domxml.createTextNode('",$value,"'));",
	               "}",

                       "if(desc.hasChildNodes())",
	                       "{",
		                 "desc.childNodes[0].nodeValue='';",
	 	                 "desc.childNodes[0].nodeValue=description;",
	                       "}", 
	                         "else",
	                       "{",
	                          "var child = desc.appendChild(domxml.createTextNode(description));",
	               "}",


                           "MLA.addCustomXMLPart(domxml.xml);",
                           "setComponentProperties();",
 
                          "}catch(err){",
                            "alert('ERROR: '+err.description);",
                            "setComponentProperties();",
                          "}",
	
                       "}",   (:end of if:)
                  "}") (:end of function:)
};
(:END Current-Document - Tags Tab - Generate Javascript Functions  :)


(:BEGIN GENERATE METADATA MAP AND TEMPLATES FROM CONFIG:)
declare function config:get-js-map(){
    let $all-controls := $config:TAGS
    let $parent-controls := ($all-controls/node()/config:workbook/config:tag, $all-controls/node()/config:worksheet/config:tag,$all-controls/node()/config:component/config:tag)
    return for $ctrl in $parent-controls
           return if(fn:empty($ctrl)) then () 
           else (fn:concat($ctrl/config:value/text(),"|", $ctrl/config:metatemplate/text()))
};

declare function config:generate-js-metadata-map-support()
{
    let $mappings := config:get-js-map()
    return fn:concat("var myparams;
                      var map = new MetadataMap();
                      MetadataMap.prototype.get = function(key)
                      {
	                return myparams[key];
                      };

                      function MetadataMap()
                      {
                        myparams = new Array();",
                        fn:string-join(
                              for $m in $mappings
                              let $props := fn:tokenize($m,"\|")
                              return   fn:concat("myparams['",$props[1],"']='", $props[2],"';"),""
                        )
                      ,"}"
            
                    )
};

declare function config:generate-js-metadata-template-func()
{
    let $templates := $config:METADATA/node()
    let $temp-cnt := fn:count($templates/config:template)
    return fn:concat("function generateTemplate(metaid){",

             fn:string-join(for $temp at $d in $templates/config:template
                            return fn:concat(
                                      if($d eq 1)then 
                                           fn:concat("if(metaid=='",$temp/@id,"'){ var v_template='") 
                                      else if($d eq $temp-cnt)then
                                           "}else{var v_template='"
                                      else fn:concat("}else if(metaid=='",$temp/@id,"'){ var v_template='"),
                                    fn:normalize-space(xdmp:quote($temp/dc:metadata)),"';")
                             ,""),"}return v_template;}")
};
(:END GENERATE METADATA MAP AND TEMPLATES FROM CONFIG:)

declare function config:generate-js-for-controls()
{
   (config:generate-js-metadata-map-support(),
    config:generate-js-metadata-template-func(), 
    config:generate-js-workbook-tag-func(),
    config:generate-js-worksheet-tag-func(),
    config:generate-js-component-tag-func()
   )
};

(:BEGIN Search Tab - Filter :)
declare function config:search-filters()
{
     let $filters := $config:SEARCH/config:searchfilters/config:searchfilter

     return <div id="searchfilter" xmlns="http://www.w3.org/1999/xhtml">
             {for $filter in $filters
              return 
               <div class="filterrow">
                   <input type="checkbox" id="{$filter/config:control-alias/text()}" />
                   <a href="#"> {$filter/config:display-label/text()}</a>
               </div>
             }
            </div>
};
(:END Search Tab - Filter :)
