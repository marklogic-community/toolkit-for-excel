/* 
Copyright 2009 Mark Logic Corporation

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

function loadXMLDoc(url) 
{
    if (window.XMLHttpRequest) {
        req = new XMLHttpRequest();
        req.onreadystatechange = processReqChange;
        req.open("GET", url, false);
        req.send(null);
        response = req.responseText;
        return response; 
    } else if (window.ActiveXObject) {
        req = new ActiveXObject("Microsoft.XMLHTTP");
        if (req) {
            req.onreadystatechange = processReqChange;
            req.open("GET", url, true);
            req.send();
        }
    }
}
function processReqChange() 
{
    // only if req shows "complete"
    if (req.readyState == 4) {
        // only if "OK"
        if (req.status == 200) {
            // ...processing statements go here...

     response = req.responseText;
        } else {
            alert("There was a problem retrieving the XML data:\n" + req.statusText);
        }
    }
}

function insertImage(picuri)
{
       var config = MLA.getConfiguration();
       var fullurl= config.url;
       var picuri = fullurl + "/officesearch/download-support.xqy?uid="+picuri;
       var msg = MLA.insertImage(picuri,"oslo","oslo");
}

function insertSlide(docuri, slideidx, retainidx)
{
       var form=document.getElementById("buttons"+retainidx);
       var type="";
       var docname="";

       for (var i = 0; i < form.searchtype.length; i++) {
         if (form.searchtype[i].checked) {
             type=form.searchtype[i].value;
	     docname=form.searchtype[i].name;
         break;
            }
       }

          if(type=="insertslide")
	  {
                  // var retain=document.getElementById("retain"+retainidx).checked;
		  var retain = "false";
                  var tmpPath = MLA.getTempPath();

                  var config = MLA.getConfiguration();
                  var fullurl= config.url;
                  var url = fullurl + "/officesearch/download-support.xqy?uid="+docuri;
      
                  var tokens = docuri.split("/");
                  var filename = tokens[tokens.length-1]; 
                  var msg = MLA.insertSlide(tmpPath, filename,slideidx, url, "oslo","oslo",retain);
	  }
	  else if(type=="opendocument")
	  { 
		  docname = docname.replace("/ppt/slides/","");
		  docname = docname.replace(/slide[0-9]+.xml/,"");
		  docname = docname.replace("_pptx_parts",".pptx");

		  var clean = docname.split("/");
		  var title = clean[clean.length-1];
		  var myref = window.location('http://localhost:8023/openbinary.xqy?url='+docuri+'&title='+title);
	  }
	  
}

function openPPTX(docuri)
{
       var tokens = docuri.split("/");
       var filename = tokens[tokens.length-1];
       var tmpPath = MLA.getTempPath(); 

       var config = MLA.getConfiguration();
       var fullurl= config.url;
       var url = fullurl + "/officesearch/download-support.xqy?uid="+docuri;

       var msg = MLA.openPPTX(tmpPath, filename, url, "oslo","oslo");
}

function openDocument(t,txt)
{
   var form=document.getElementById("buttons"+t);
   var type="";
   var docname="";

   for (var i = 0; i < form.searchtype.length; i++) {
      if (form.searchtype[i].checked) {
          type=form.searchtype[i].value;
	  docname=form.searchtype[i].name
      break;
      }
   }

          if(type=="inserttext")
	  {
		  MLA.insertText(txt);
	  }
	  else if(type=="opendocument")
	  {
		  //title for word doc
		  docname = docname.replace("/word/document.xml","");
		  docname = docname.replace("_docx_parts",".docx");

		  //title for xl doc
		  docname = docname.replace("/xl/worksheets/","");
		  docname = docname.replace(/sheet[0-9]+.xml/,"");
		  docname = docname.replace("_xlsx_parts",".xlsx");
		  
		  var clean = docname.split("/");
		  var title = clean[clean.length-1];
		  var myref = window.location('http://localhost:8023/openbinary.xqy?url='+docname+'&title='+title);

	  }
	  else if(type=="embeddocument")
	  {
		  //title for word doc
		  docname = docname.replace("/word/document.xml","");
		  docname = docname.replace("_docx_parts",".docx");


		  //title for xl doc
		  docname = docname.replace("/xl/worksheets/","");
		  docname = docname.replace(/sheet[0-9]+.xml/,"");
		  docname = docname.replace("_xlsx_parts",".xlsx");

		  var clean = docname.split("/");
		  var title = clean[clean.length-1];


		  var tmpPath = MLA.getTempPath(); 

                  var config = MLA.getConfiguration();
                  var fullurl= config.url;
                  var url = fullurl + "/officesearch/download-support.xqy?uid="+docname;
		  //alert("tmppath: "+tmpPath+"\n  url: "+url+ "\n   title: "+title);

                  var msg = MLA.embedOLE(tmpPath, title, url, "oslo","oslo");
	  }

}
