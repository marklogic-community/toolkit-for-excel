xquery version "1.0-ml";
(:
Copyright 2008-2010 Mark Logic Corporation

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

module namespace config="http://marklogic.com/config";
declare namespace dc="http://purl.org/dc/elements/1.1/";

declare variable $config:CONTROLS := xdmp:document-get("Docs/Author/config/controls.xml");
declare variable $config:BOILERPLATE := xdmp:document-get("Docs/Author/config/boilerplate.xml");
declare variable $config:METADATA := xdmp:document-get("Docs/Author/config/metadata.xml");
declare variable $config:SEARCH := xdmp:document-get("Docs/Author/config/search.xml");
declare variable $config:COMPARE := xdmp:document-get("Docs/Author/config/compare.xml");

(:BEGIN Current-Document - Controls Tab Display:)
declare function config:textctrl-sections()
{
    let $text-sections := $config:CONTROLS/node()/config:richtext/config:section
    for $t at $d in $text-sections
    let $func := fn:concat("txtSectionFunc",$d,"()")
    return <li>
             <a href="#" onmouseup="blurSelected(this)" onclick={$func}>{$t/config:title/text()}</a>
           </li>

};

declare function config:textctrl-inline()
{
    let $text-inline := $config:CONTROLS/node()/config:richtext/config:inline
    for $t at $d in $text-inline
    let $func := fn:concat("txtInlineFunc",$d,"()")
    return <li>
             <a href="#" onmouseup="blurSelected(this)" onclick={$func}>{$t/config:title/text()}</a>
           </li>

};

declare function config:picctrl-inline()
{
    let $img-inline := $config:CONTROLS/node()/config:image/config:inline
    for $t at $d in $img-inline
    let $func := fn:concat("picInlineFunc",$d,"()")
    return <li>
             <a href="#" onmouseup="blurSelected(this)" onclick={$func}>{$t/config:title/text()}</a>
           </li>
};

declare function config:calctrl-inline()
{
    let $cal-inline := $config:CONTROLS/node()/config:calendar/config:inline
    for $t at $d in $cal-inline
    let $func := fn:concat("calInlineFunc",$d,"()")
    return <li>
             <a href="#" onmouseup="blurSelected(this)" onclick={$func}>{$t/config:title/text()}</a>
           </li>
};

declare function config:dropctrl-inline()
{
    let $drop-inline := $config:CONTROLS/node()/config:dropdown/config:inline
    for $t at $d in $drop-inline
    let $func := fn:concat("dropInlineFunc",$d,"()")
    return <li>
             <a href="#" onmouseup="blurSelected(this)" onclick={$func}>{$t/config:title/text()}</a>
           </li>
};

declare function config:comboctrl-inline()
{
    let $combo-inline := $config:CONTROLS/node()/config:combo/config:inline
    for $t at $d in $combo-inline
    let $func := fn:concat("comboInlineFunc",$d,"()")
    return <li>
             <a href="#" onmouseup="blurSelected(this)" onclick={$func}>{$t/config:title/text()}</a>
           </li>
};
(:END Current-Document - Controls Tab Display:)

(:BEGIN Current-Document - Controls Tab - Generate Javascript Functions  :)
declare function config:control-type($node as node()) as xs:string?
{
  if(fn:node-name($node) eq fn:QName("http://marklogic.com/config", "richtext")) then
      "wdContentControlRichText"
  else if(fn:node-name($node) eq fn:QName("http://marklogic.com/config", "image")) then
      "wdContentControlPicture"
  else if(fn:node-name($node) eq fn:QName("http://marklogic.com/config", "calendar")) then
       "wdContentControlDate"
  else if(fn:node-name($node) eq fn:QName("http://marklogic.com/config", "dropdown")) then
       "wdContentControlDropdownList"
  else if(fn:node-name($node) eq fn:QName("http://marklogic.com/config", "combo")) then
       "wdContentControlComboBox"
  else ""

};

declare function config:generate-js-for-child-ctrl($children as node()*, $idx as xs:string?) as xs:string
{fn:string-join(
  for $child at $d in $children/child::*
  let $type :=  config:control-type($child)
  let $title := $child/config:title/text()
  let $ph-text := if(fn:empty($child/config:placeholdertext/text())) then ()
                    else 
                          $child/config:placeholdertext/text()
  let $subchildren := $child/config:children
  let $newline := $child/config:newline/text()
  let $newidx := fn:concat("",$idx,$d)
  let $parent-id := if(fn:empty($idx)) then "ccid" else fn:concat("childId",$newidx)
  return fn:concat("var childId",$newidx," = MLA.addContentControl('','",$title,"','", $type,"','", $newline,"',",$parent-id,");",
                          if(fn:not(fn:empty($ph-text))) then 
                                (fn:concat("MLA.insertContentControlText(childId",$newidx,",'",$ph-text,"');",
                                            "MLA.setContentControlPlaceholderText(childId",$d,",'",$ph-text,"');"))  else "",
                                 config:generate-js-for-child-ctrl($subchildren,$newidx)
                          
               
                ) , "")
  
};

declare function config:generate-js-section-text()
{
    let $text-section-ctrls := $config:CONTROLS/node()/config:richtext/config:section
    for $tc at $d in $text-section-ctrls
    let $title := $tc/config:title/text()
    let $ph-text := if(fn:empty($tc/config:placeholdertext/text())) then ()
                    else 
                          $tc/config:placeholdertext/text()    
    let $children := $tc/config:children
    let $newline := "true"
    let $type := "wdContentControlRichText"
    return fn:concat("function txtSectionFunc",$d,"(){ 
                      var ccid  = MLA.addContentControl('','",$title,"','", $type,"','", $newline,"', '');", 
                                 "MLA.setContentControlTag(ccid,ccid);",

                   if(fn:not(fn:empty($ph-text))) then 
                                fn:concat("MLA.setContentControlPlaceholderText(ccid,'",$ph-text,"');") else "",
                    config:generate-js-for-child-ctrl($children,()), 
                  "}")
};

declare function config:generate-js-inline-text()
{

    let $text-inline-ctrls :=$config:CONTROLS/node()/config:richtext/config:inline
    for $tc at $d in $text-inline-ctrls
    let $title := $tc/config:title/text()
    let $ph-text := if(fn:empty($tc/config:placeholdertext/text())) then ()
                    else 
                          $tc/config:placeholdertext/text()    
    let $newline := "false"
    let $type := "wdContentControlRichText"
    return fn:concat("function txtInlineFunc",$d,"(){ 
                      var ccid  = MLA.addContentControl('','",$title,"','", $type,"','", $newline,"', '');", 
                                 "MLA.setContentControlTag(ccid,ccid);",

                   if(fn:not(fn:empty($ph-text))) then 
                                fn:concat("MLA.setContentControlPlaceholderText(ccid,'",$ph-text,"');") else "",
                  "}")
};

declare function config:generate-js-inline-pic()
{
    let $pic-inline-ctrls := $config:CONTROLS/node()/config:image/config:inline
    for $pc at $d in $pic-inline-ctrls
    let $title := $pc/config:title/text()
    let $newline := "false"
    let $type := "wdContentControlPicture"
    return fn:concat("function picInlineFunc",$d,"(){ 
                      var ccid  = MLA.addContentControl('','",$title,"','", $type,"','", $newline,"', '');", 
                                 "MLA.setContentControlTag(ccid,ccid);",
                  "}")
};

declare function config:generate-js-inline-cal()
{
    let $cal-inline-ctrls := $config:CONTROLS/node()/config:calendar/config:inline
    for $cc at $d in $cal-inline-ctrls
    let $title := $cc/config:title/text()
    let $ph-text := if(fn:empty($cc/config:placeholdertext/text())) then ()
                    else 
                          $cc/config:placeholdertext/text()    
    let $newline := "false"
    let $type := "wdContentControlDate"
    return fn:concat("function calInlineFunc",$d,"(){ 
                      var ccid  = MLA.addContentControl('','",$title,"','", $type,"','", $newline,"', '');", 
                                 "MLA.setContentControlTag(ccid,ccid);",
                  if(fn:not(fn:empty($ph-text))) then 
                                fn:concat("MLA.setContentControlPlaceholderText(ccid,'",$ph-text,"');") else "",
                  "}")
};

declare function config:generate-js-inline-drop()
{
    let $drop-inline-ctrls := $config:CONTROLS/node()/config:dropdown/config:inline
    for $dc at $d in $drop-inline-ctrls
    let $title := $dc/config:title/text()
    let $ph-text := if(fn:empty($dc/config:placeholdertext/text())) then ()
                    else 
                          $dc/config:placeholdertext/text()
    let $le := $dc/config:entry    
    let $newline := "false"
    let $type := "wdContentControlDropdownList"
    return fn:concat("function dropInlineFunc",$d,"(){ 
                      var ccid  = MLA.addContentControl('','",$title,"','", $type,"','", $newline,"', '');", 
                                 "MLA.setContentControlTag(ccid,ccid);",
                  if(fn:not(fn:empty($ph-text))) then 
                                fn:concat("MLA.setContentControlPlaceholderText(ccid,'",$ph-text,"');") else "",
                  fn:string-join(for $l in $le
                 return fn:concat("MLA.addContentControlDropDownListEntries(ccid,'",$l/config:text/text(),"','",$l/config:value/text(),"','0');" ),""),
                  "}")
};

declare function config:generate-js-inline-combo()
{
    let $combo-inline-ctrls := $config:CONTROLS/node()/config:combo/config:inline
    for $dc at $d in $combo-inline-ctrls
    let $title := $dc/config:title/text()
    let $ph-text := if(fn:empty($dc/config:placeholdertext/text())) then ()
                    else 
                          $dc/config:placeholdertext/text()
    let $list-entries := $dc/config:entry    
    let $newline := "false"
    let $type := "wdContentControlComboBox"
    return fn:concat("function comboInlineFunc",$d,"(){ 
                      var ccid  = MLA.addContentControl('','",$title,"','", $type,"','", $newline,"', '');", 
                                 "MLA.setContentControlTag(ccid,ccid);",
                  if(fn:not(fn:empty($ph-text))) then 
                                fn:concat("MLA.setContentControlPlaceholderText(ccid,'",$ph-text,"');") else "",
                  fn:string-join(for $l in $list-entries
                 return fn:concat("MLA.addContentControlDropDownListEntries(ccid,'",$l/config:text/text(),"','",$l/config:value/text(),"','0');" ),""),
                  "}")
};

(:BEGIN GENERATE METADATA MAP AND TEMPLATES FROM CONFIG:)
declare function config:get-map-subs($node as node()*) as xs:string*
{
    for $n in $node 
    return  if(fn:empty($n)) then () 
            else (fn:concat($n/config:title/text(),"|", $n/config:metatemplate/text()),
                  config:get-map-subs($n/child::*//config:children/child::*))       
};

declare function config:get-js-map(){
    let $all-controls := $config:CONTROLS
    let $parent-controls := ($all-controls/child::*/child::*/config:section, $all-controls//config:inline)
    return for $ctrl in $parent-controls
           return if(fn:empty($ctrl)) then () 
           else (fn:concat($ctrl/config:title/text(),"|", $ctrl/config:metatemplate/text()),
                 config:get-map-subs( $ctrl/config:children/child::* ))
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
    config:generate-js-section-text(),
    config:generate-js-inline-text(), 
    config:generate-js-inline-pic(),
    config:generate-js-inline-cal(),
    config:generate-js-inline-drop(),
    config:generate-js-inline-combo()
   )
};
(:END Current-Document - Controls Tab - Generate Javascript Functions :)

(:BEGIN Current-Document - Snippets Tab :)
declare function config:snippets()
{
     let $doc := $config:BOILERPLATE 
     let $bps := $doc/config:boilerplates/config:boilerplate
     return
     <ul xmlns="http://www.w3.org/1999/xhtml" class="buttongroup">
     {
       for $bp in $bps
       let $uri := $bp/config:document-uri/text()
       let $icon := $bp/config:icon/text()
       let $display-image := if($icon eq "word") then
                               "lrgbtn lrgbtnword"
                             else if($icon eq "image") then
                               "lrgbtn lrgbtnimage"
                             else 
                               "lrgbtn lrgbtnchart"
                              
       return   
         (<li>
           <a href="#" class="{$display-image}" onmouseup="blurSelected(this)" onclick="boilerplateinsert('{$uri}')" alt="{$uri}" title="{$uri}">
             {$bp/config:document-label/text()}
           </a>
         </li>, 
         <br clear="all"/>
         )  
      }
      </ul>
};
(:END Current-Document - Snippets Tab :)

(:BEGIN Search Tab - Filter :)
declare function config:search-filters()
{
     let $filters := $config:SEARCH/config:searchfilters/config:searchfilter

     return <div id="searchfilter" xmlns="http://www.w3.org/1999/xhtml">
             {for $filter in $filters
              return 
               <div class="filterrow">
                   <input type="checkbox" id="{$filter/config:controlalias/text()}" />
                   <a href="#"> {$filter/config:displaylabel/text()}</a>
               </div>
             }
            </div>
};

(:END Search Tab - Filter :)

(:BEGIN Compare Tab - Filter :)
declare function config:compare-filters()
{
     let $filters := $config:COMPARE/config:comparefilters/config:comparefilter
     return  <ul id="ulmenu" onClick="displayLayer('ulmenu')">
             {
              for $f at $d in $filters
              return <li onClick="setSelected({$d})">
                         <a href="#" class="menu" id="{fn:concat('select',$d)}" name="{$f/config:value/text()}">{$f/config:displaylabel/text()}</a>
                     </li>
             }</ul>


};
(:END Compare Tab - Filter :)
