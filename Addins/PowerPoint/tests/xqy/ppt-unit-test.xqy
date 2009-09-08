xquery version "1.0-ml";
(: Copyright 2009 Mark Logic Corporation

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

declare namespace html = "http://www.w3.org/1999/xhtml";
declare namespace test = "http://test";
declare namespace p= "http://schemas.openxmlformats.org/presentationml/2006/main";
declare namespace r="http://schemas.openxmlformats.org/officeDocument/2006/relationships";
declare namespace rel="http://schemas.openxmlformats.org/package/2006/relationships";

import module namespace ppt= "http://marklogic.com/openxml/powerpoint"
                          at "/MarkLogic/openxml/presentation-ml-support.xqy";

(: Testing  ppt:insert-slide for different presos,with different combinations, 
   at different start positions. Use MarkLogic_PowerPointAddin_Test.exe to loop 
   through and open/close presentations, log errors 
:)

declare function ppt:test-one()
{
let $t-pres:="/two_pptx_parts/"       (:target presentation:)
let $s-pres:="/one_pptx_parts/"       (:source presentation:)
let $s-idx := 1             
let $start-idx := 2       
let $pptx-map := ppt:package-map($t-pres)
let $new-map  := ppt:insert-slide($pptx-map, $s-pres, $s-idx, $start-idx)
let $pptx-pkg := ppt:package-map-zip($new-map)  
return ($new-map, xdmp:save("C:\unitTestAddin\pptx\VALIDATE_ONE.pptx",$pptx-pkg))  
};

declare function ppt:test-two()
{
let $t-pres:="/two_pptx_parts/"       (:target presentation:)
let $s-pres:="/one_pptx_parts/"       (:source presentation:)
let $s-pres3 := "/testOne_pptx_parts/"
let $s-idx := 1             (:index of slide in source to copy to target :)
let $s-idx2 := 2
let $s-idx3 := 3
let $start-idx := 3         (:insertion index of target presentation :)
let $pptx-map := ppt:package-map($t-pres)
let $new-map  := ppt:insert-slide($pptx-map, ($s-pres,$s-pres,$s-pres3), ($s-idx,$s-idx2, $s-idx3), $start-idx)
let $pptx-pkg := ppt:package-map-zip($new-map)  
return ($new-map   , xdmp:save("C:\unitTestAddin\pptx\VALIDATE_TWO.pptx",$pptx-pkg))  
};

declare function ppt:test-three()
{
let $t-pres:="/two_pptx_parts/"       (:target presentation:)
let $s-pres:="/one_pptx_parts/"       (:source presentation:)
let $s-pres3 := "/testOne_pptx_parts/"
let $s-idx := 1             (:index of slide in source to copy to target :)
let $s-idx2 := 2
let $s-idx3 := 3
let $start-idx := 1         (:insertion index of target presentation :)
let $pptx-map := ppt:package-map($t-pres)
let $new-map  := ppt:insert-slide($pptx-map, ($s-pres,$s-pres,$s-pres3), ($s-idx,$s-idx2, $s-idx3), $start-idx)
let $pptx-pkg := ppt:package-map-zip($new-map)  
return ($new-map   , xdmp:save("C:\unitTestAddin\pptx\VALIDATE_THREE.pptx",$pptx-pkg))
};

declare function ppt:test-four()
{
let $t-pres:="/two_pptx_parts/"       (:target presentation:)
let $s-pres:="/one_pptx_parts/"       (:source presentation:)
let $s-pres3 := "/testOne_pptx_parts/"
let $s-idx := 1             (:index of slide in source to copy to target :)
let $s-idx2 := 2
let $s-idx3 := 3
let $start-idx := 2         (:insertion index of target presentation :)
let $pptx-map := ppt:package-map($t-pres)
let $new-map  := ppt:insert-slide($pptx-map, ($s-pres,$s-pres,$s-pres3), ($s-idx,$s-idx2, $s-idx3), $start-idx)
let $pptx-pkg := ppt:package-map-zip($new-map)  
return ($new-map   , xdmp:save("C:\unitTestAddin\pptx\VALIDATE_FOUR.pptx",$pptx-pkg)) 
};

declare function ppt:test-five()
{
let $t-pres:="/Aven_MarkLogicUserConference2009Exceling_pptx_parts/" 
let $s-pres:="/two_pptx_parts/"           
let $s-pres3 := "/testOne_pptx_parts/"
let $s-idx := 1             (:index of slide in source to copy to target :)
let $s-idx2 := 2
let $s-idx3 := 3
let $start-idx := 2         (:insertion index of target presentation :)
let $pptx-map := ppt:package-map($t-pres)
let $new-map  := ppt:insert-slide($pptx-map, ($s-pres,$s-pres,$s-pres3), ($s-idx,$s-idx2, $s-idx3), $start-idx)
let $pptx-pkg := ppt:package-map-zip($new-map)  
return ($new-map   , xdmp:save("C:\unitTestAddin\pptx\VALIDATE_FIVE.pptx",$pptx-pkg)) 
};

declare function ppt:test-six()
{
let $t-pres:="/Aven_MarkLogicUserConference2009Exceling_pptx_parts/"       
let $s-pres:="/one_pptx_parts/"  
let $s-pres3 := "/testOne_pptx_parts/"
let $s-idx := 1             (:index of slide in source to copy to target :)
let $s-idx2 := 2
let $s-idx3 := 3
let $start-idx := 2         (:insertion index of target presentation :)
let $pptx-map := ppt:package-map($t-pres)
let $new-map  := ppt:insert-slide($pptx-map, ($s-pres,$s-pres,$s-pres3), ($s-idx,$s-idx2, $s-idx3), $start-idx)
let $pptx-pkg := ppt:package-map-zip($new-map)  
return ($new-map   , xdmp:save("C:\unitTestAddin\pptx\VALIDATE_SIX.pptx",$pptx-pkg))
};

declare function ppt:test-seven()
{
let $presentation-dir:="/testOne_pptx_parts/" 
let $uris := ppt:directory-uris($presentation-dir) 
return xdmp:save("C:\unitTestAddin\outputs\testOne.xml",ppt:package-make($presentation-dir, $uris))
};

declare function ppt:test-eight()
{
let $presentation-dir:="/testOne_pptx_parts/"
let $map := ppt:package-map($presentation-dir)
return xdmp:save("C:\unitTestAddin\outputs\map.xml",text{$map}) 
};

ppt:test-one(),
ppt:test-two(),
ppt:test-three(),
ppt:test-four(),
ppt:test-five(),
ppt:test-six(),
(:saves pkg.xml to /outputs, test for open in ppt:)
ppt:test-seven(),
(:saves map.xml to /outputs :)
ppt:test-eight()


