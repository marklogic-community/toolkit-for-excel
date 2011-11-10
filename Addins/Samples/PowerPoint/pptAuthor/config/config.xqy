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

module namespace config="http://marklogic.com/toolkit/powerpoint/author/config";
declare namespace dc="http://purl.org/dc/elements/1.1/";


declare variable $config:CONFIG-PATH := "http://localhost:8023/pptAuthor/config/";
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
declare function config:presentation-tags()
{
    let $img-inline := $config:TAGS/node()/config:presentation/config:tag
    for $t at $d in $img-inline
    let $func := fn:concat("presentationTagFunc",$d,"()")
    return <li>
             <a href="#" onmouseup="blurSelected(this)" onclick="{$func}">{$t/config:name/text()}</a>
           </li>


};

declare function config:slide-tags()
{
    let $img-inline := $config:TAGS/node()/config:slide/config:tag
    for $t at $d in $img-inline
    let $func := fn:concat("slideTagFunc",$d,"()")
    return <li>
             <a href="#" onmouseup="blurSelected(this)" onclick="{$func}">{$t/config:name/text()}</a>
           </li>


};

declare function config:shape-tags()
{
    let $img-inline := $config:TAGS/node()/config:shape/config:tag
    for $t at $d in $img-inline
    let $func := fn:concat("shapeTagFunc",$d,"()")
    return <li>
             <a href="#" onmouseup="blurSelected(this)" onclick="{$func}">{$t/config:name/text()}</a>
           </li>


};
(:END Current-Document - TAGs Tab Display:)

(:BEGIN Current-Document - Tags Tab - Generate Javascript Functions  :)
declare function config:generate-js-presentation-tag-func()
{
    let $preso-tags := $config:TAGS/node()/config:presentation/config:tag
    for $pt at $d in $preso-tags
    let $name := $pt/config:name/text() (:display label:)
    let $value := $pt/config:value/text()  (:used for name:)
    return fn:concat("function presentationTagFunc",$d,"(){",
                       "var tagId = randomId();",
                       "if(!(checkPresentationTags('",$value,"'))){;",
                       "var msg = MLA.addPresentationTag('",$value,"',tagId);",

                       "var presoTags = MLA.getPresentationTags();",
                       "var myTagsString =getJsonString(presoTags);",

                       "setPresentationProperties();",

                       "var stringxml = MLA.unescapeXMLCharEntities(generateTemplate(map.get('",$value,"')));",
                       "var domxml = MLA.createXMLDOM(stringxml);",
                       "var id = domxml.getElementsByTagName('dc:identifier')[0];",
                       "var jsonStore = domxml.getElementsByTagName('dc:description')[0];",
                       

	                   "if(id.hasChildNodes())",
	                   "{",
		              "id.childNodes[0].nodeValue='';",
	 	              "id.childNodes[0].nodeValue=tagId;",
	                   "}", 
	                   "else",
	                   "{",
	                      "var child = id.appendChild(domxml.createTextNode(tagId));",
	                   "}",

                           "if(jsonStore.hasChildNodes())",
	                       "{",
		                 "jsonStore.childNodes[0].nodeValue='';",
	 	                 "jsonStore.childNodes[0].nodeValue=myTagsString;",
	                       "}", 
	                         "else",
	                       "{",
	                          "var child = jsonStore.appendChild(domxml.createTextNode(myTagsString));",
	                    "}",

                           "MLA.addCustomXMLPart(domxml.xml);",

                       (:"alert(stringxml+'FOO'+domxml.xml);:) 
                           "}",
	
                  "}") 
};

declare function  config:generate-js-slide-tag-func()
{
    let $slide-tags := $config:TAGS/node()/config:slide/config:tag
    for $st at $d in $slide-tags
    let $name := $st/config:name/text()
    let $value := $st/config:value/text()
    return fn:concat("function slideTagFunc",$d,"(){",
                       "var tagId = randomId();",
                       "if(!(checkSlideTags('",$value,"'))){",
                       "var idx = MLA.getSlideIndex();",
                       "var msg = MLA.addSlideTag(idx,'",$value,"',tagId);",
                       "var slideTags = MLA.getSlideTags(idx);",
                       "var myTagsString =getJsonString(slideTags);",

                       "setSlideProperties();",

                       "var stringxml = MLA.unescapeXMLCharEntities(generateTemplate(map.get('",$value,"')));",
                       "var domxml = MLA.createXMLDOM(stringxml);",
                       "var id = domxml.getElementsByTagName('dc:identifier')[0];",
                       "var jsonStore = domxml.getElementsByTagName('dc:description')[0];",
                       

	                   "if(id.hasChildNodes())",
	                   "{",
		              "id.childNodes[0].nodeValue='';",
	 	              "id.childNodes[0].nodeValue=tagId;",
	                   "}", 
	                   "else",
	                   "{",
	                      "var child = id.appendChild(domxml.createTextNode(tagId));",
	                   "}",

                           "if(jsonStore.hasChildNodes())",
	                       "{",
		                 "jsonStore.childNodes[0].nodeValue='';",
	 	                 "jsonStore.childNodes[0].nodeValue=myTagsString;",
	                       "}", 
	                         "else",
	                       "{",
	                          "var child = jsonStore.appendChild(domxml.createTextNode(myTagsString));",
	                   "}",

                           "MLA.addCustomXMLPart(domxml.xml);",
                           (: need to serialize and save tags in description :)
                           (:"alert(stringxml+'FOO'+domxml.xml);:) 
                           "}",
	
                  "}") 
};

declare function  config:generate-js-shape-tag-func()
{
    let $shape-tags := $config:TAGS/node()/config:shape/config:tag
    for $st at $d in $shape-tags
    let $name := $st/config:name/text()
    let $value := $st/config:value/text()
    return fn:concat("function shapeTagFunc",$d,"(){",
                       "var tagId = randomId();",
                       "if(!(checkShapeTags('",$value,"'))){",
                       (:"var idx = MLA.getSlideIndex();",:)
                          "try{", 
                               "var shapename = MLA.getShapeRangeName();",
                               "var slideindex = MLA.getSlideIndex();",
                                                                      
                               "var msg = MLA.addShapeTag(slideindex, shapename,'",$value,"',tagId);",

                               "var shapeRange = MLA.getShapeRangeView(slideindex, shapename);",
                               "setShapeProperties();",

                               "var stringxml = MLA.unescapeXMLCharEntities(generateTemplate(map.get('",$value,"')));",
                               "var domxml = MLA.createXMLDOM(stringxml);",

                               "var id = domxml.getElementsByTagName('dc:identifier')[0];",
                               "var jsonStore = domxml.getElementsByTagName('dc:description')[0];",
                                           
                               "var myShapeString =getJsonString(shapeRange);", 

	                       "if(id.hasChildNodes())",
	                       "{",
		                 "id.childNodes[0].nodeValue='';",
	 	                 "id.childNodes[0].nodeValue=tagId;",
	                       "}", 
	                         "else",
	                       "{",
	                          "var child = id.appendChild(domxml.createTextNode(tagId));",
	                       "}",

	                       "if(jsonStore.hasChildNodes())",
	                       "{",
		                 "jsonStore.childNodes[0].nodeValue='';",
	 	                 "jsonStore.childNodes[0].nodeValue=myShapeString;",
	                       "}", 
	                         "else",
	                       "{",
	                          "var child = jsonStore.appendChild(domxml.createTextNode(myShapeString));",
	                       "}",

                                
                                  "MLA.addCustomXMLPart(domxml.xml);",

                          "}catch(err){",
                            "setShapeProperties();",
                          "}",
	
                       "}",  (:end of if:)
                  "}") (:end of function:)
};
(:END Current-Document - Tags Tab - Generate Javascript Functions  :)


(:BEGIN GENERATE METADATA MAP AND TEMPLATES FROM CONFIG:)
declare function config:get-js-map(){
    let $all-controls := $config:TAGS
    let $parent-controls := ($all-controls/node()/config:presentation/config:tag, $all-controls/node()/config:slide/config:tag,$all-controls/node()/config:shape/config:tag)
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
    config:generate-js-presentation-tag-func(),
    config:generate-js-slide-tag-func(),
    config:generate-js-shape-tag-func()
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
