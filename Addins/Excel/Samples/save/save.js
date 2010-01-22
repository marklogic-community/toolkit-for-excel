/*
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
*/
function saveXlsxToML()
{
       var ele = document.getElementById("ML-Save");
       var doctitle = ele.value;
       if(doctitle=="")
       {
	   doctitle="Default.xlsx";
       }
       else
       {
	   doctitle=doctitle+".xlsx";
       }

       var tmpPath = MLA.getTempPath(); 

       var config = MLA.getConfiguration();
       var fullurl= config.url;
       var url = fullurl + "/save/upload2.xqy?uid="+doctitle;

       var saveas = tmpPath+doctitle;

       var msg = MLA.saveActiveWorkbook(tmpPath, doctitle, url, "zeke","zeke");

       if(msg=="")
	       alert("workbook:" + doctitle + " saved.");
}
	
