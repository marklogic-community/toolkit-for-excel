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
       //alert("config url"+fullurl);
       var picuri = fullurl + "/search/download-support.xqy?uid="+picuri;
       var msg = MLA.insertImage(picuri,"oslo","oslo");
}

//function copyPasteSlideToActive(docuri)
function copyPasteSlideToActive(docuri, slideidx)
{
	//alert("docuri "+docuri+" slideidx"+slideidx);

       // alert("here");
  /*      alert("docuri for copyslidetoactive():"+ docuri);
       var tokens = docuri.split("/");
       var filename = tokens[tokens.length-1];

       //use filename to get slide number for now
       //get .pptx name from docuri for now
       //until we embed in xml somewhere (properties?)

       //alert("filename "+filename);
       var tmpfilename = filename.replace(".GIF","");
       tmpfilename = tmpfilename.replace("Slide","");

       //var slideidx = parseInt(tmpfilename);
       var slideidx = tmpfilename;
       //alert("tmpfilename "+tmpfilename + "slideidx: "+slideidx);
       var idx = docuri.indexOf("_GIF");
       var tmpuri1 = docuri.substring(0,idx);
       
       var tmpuri2 = docuri.substring(1,idx);
       var newuri = tmpuri1+".pptx";
       var newfilename = tmpuri2+".pptx";
       

       //alert("newuri: "+newuri);

       //alert("filename"+filename);
       //*/
       var tmpPath = MLA.getTempPath();
       //alert("here2 "+tmpPath) 

       var config = MLA.getConfiguration();
       var fullurl= config.url;
       var url = fullurl + "/search/download-support.xqy?uid="+docuri;

       var filename = docuri.substring(1,docuri.length);
       //var url = fullurl + "/search/download-support.xqy?uid="+newuri;

       //alert("fullurl"+fullurl);
       //alert("URL: "+url);
	//window.external.openXlsx(filename,docuri);
       var msg = MLA.copyPasteSlideToActive(tmpPath, filename,slideidx, url, "oslo","oslo");
     // var msg = window.external.OpenXlsx(tmpPath, docuri, url, "zeke","zeke");
}

function insertImage(picuri)
{
       var config = MLA.getConfiguration();
       var fullurl= config.url;
       //alert("config url"+fullurl);
       var picuri = fullurl + "/search/download-support.xqy?uid="+picuri;
       var msg = MLA.insertImage(picuri,"oslo","oslo");
}

//function copyPasteSlideToActive(docuri)
function copyPasteSlideToActive(docuri, slideidx,retainidx)
{
var retain=document.getElementById("retain"+retainidx).checked;
//alert("value is"+x + " retainidx = "+retainidx);
	//alert("docuri "+docuri+" slideidx"+slideidx);

       // alert("here");
  /*      alert("docuri for copyslidetoactive():"+ docuri);
       var tokens = docuri.split("/");
       var filename = tokens[tokens.length-1];

       //use filename to get slide number for now
       //get .pptx name from docuri for now
       //until we embed in xml somewhere (properties?)

       //alert("filename "+filename);
       var tmpfilename = filename.replace(".GIF","");
       tmpfilename = tmpfilename.replace("Slide","");

       //var slideidx = parseInt(tmpfilename);
       var slideidx = tmpfilename;
       //alert("tmpfilename "+tmpfilename + "slideidx: "+slideidx);
       var idx = docuri.indexOf("_GIF");
       var tmpuri1 = docuri.substring(0,idx);
       
       var tmpuri2 = docuri.substring(1,idx);
       var newuri = tmpuri1+".pptx";
       var newfilename = tmpuri2+".pptx";
       

       //alert("newuri: "+newuri);

       //alert("filename"+filename);
       //*/
       var tmpPath = MLA.getTempPath();
       //alert("here2 "+tmpPath) 

       var config = MLA.getConfiguration();
       var fullurl= config.url;
       var url = fullurl + "/search/download-support.xqy?uid="+docuri;

       var filename = docuri.substring(1,docuri.length);
       //var url = fullurl + "/search/download-support.xqy?uid="+newuri;

       //alert("fullurl"+fullurl);
       //alert("URL: "+url);
	//window.external.openXlsx(filename,docuri);
       var msg = MLA.copyPasteSlideToActive(tmpPath, filename,slideidx, url, "oslo","oslo",retain);
     // var msg = window.external.OpenXlsx(tmpPath, docuri, url, "zeke","zeke");
}

function openPPTX(docuri)
{
//	 alert("docuri for testOpen():"+ docuri);
        var tokens = docuri.split("/");
	var filename = tokens[tokens.length-1];
         //alert("filename"+filename);
       var tmpPath = MLA.getTempPath(); 

       var config = MLA.getConfiguration();
       var fullurl= config.url;
       var url = fullurl + "/search/download-support.xqy?uid="+docuri;
     //  alert("URL: "+url);
	//window.external.openXlsx(filename,docuri);
     var msg = window.external.openPPTX(tmpPath, docuri, url, "oslo","oslo");
     // var msg = window.external.OpenXlsx(tmpPath, docuri, url, "zeke","zeke");
}
function openWord(t,txt)
{
//alert("TEST"+t+txt);
   var form=document.getElementById("buttons"+t);
   var type="";
   var docname="";
   //var form = document.forms[0];
   for (var i = 0; i < form.searchtype.length; i++) {
      if (form.searchtype[i].checked) {
          type=form.searchtype[i].value;
	  docname=form.searchtype[i].name
      break;
      }
   }

          if(type=="inserttext")
	  {
		  //alert("true"+docname);
		  window.external.insertText(txt);
	  }else if(type=="opendocument")
	  {
		  //alert(docname);
		  docname = docname.replace("/word/document.xml","");
		  docname = docname.replace("_docx_parts",".docx");
		  docname = docname.replace("/xl/worksheets/","");
		  docname = docname.replace(/sheet[0-9]+.xml/,"");
		  docname = docname.replace("_xlsx_parts",".xlsx");
		  //alert(docname);
		  
		  var clean = docname.split("/");
		  var title = clean[clean.length-1];
		  var myref = window.location('http://localhost:8023/openbinary.xqy?url='+docname+'&title='+title);

	  }
	  else
	  {
		  docname = docname.replace("/xl/worksheets/","");
		  docname = docname.replace(/sheet[0-9]+.xml/,"");
		  docname = docname.replace("_xlsx_parts",".xlsx");
		         var tmpPath = MLA.getTempPath(); 

                  var config = MLA.getConfiguration();
                  var fullurl= config.url;
                  var url = fullurl + "/officesearch/download-support.xqy?uid="+docname;
		  alert("fullurl"+url);
                  var msg = window.external.embedXLSX(tmpPath, docname, url, "oslo","oslo")
		  //window.external.embedXLSX();
	          alert("foo: "+docname);
	  }

}

function test()
{
	alert("testing");
}
