/* 
Copyright 2009-2011 MarkLogic Corporation

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

function insertImage(picuri)
{
       var config = MLA.getConfiguration();
       var fullurl= config.url;
       var picuri = fullurl + "/search/download-support.xqy?uid="+picuri;
       var msg = MLA.insertImage(picuri,"uname","pwd");
}

function insertSlide(docuri, slideidx, retainidx)
{

       var retain=document.getElementById("retain"+retainidx).checked;
       var tmpPath = MLA.getTempPath();

       var config = MLA.getConfiguration();
       var fullurl= config.url;
       var url = fullurl + "/search/download-support.xqy?uid="+docuri;
      
       var tokens = docuri.split("/");
       var filename = tokens[tokens.length-1]; 
       var msg = MLA.insertSlide(tmpPath, filename,slideidx, url, "uname","pwd",retain);
}

function openPPTX(docuri)
{
       var tokens = docuri.split("/");
       var filename = tokens[tokens.length-1];
       var tmpPath = MLA.getTempPath(); 

       var config = MLA.getConfiguration();
       var fullurl= config.url;
       var url = fullurl + "/search/download-support.xqy?uid="+docuri;

       var msg = MLA.openPPTX(tmpPath, filename, url, "uname","pwd");
}

