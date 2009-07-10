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

function copyPasteSlideToActive(docuri, slideidx,retainidx)
{
       //alert("in this function"+retainidx);

       var retain=document.getElementById("retain"+retainidx).checked;
       var tmpPath = MLA.getTempPath();

       var config = MLA.getConfiguration();
       var fullurl= config.url;
       var url = fullurl + "/officesearch/download-support.xqy?uid="+docuri;
       
       var filename = docuri.substring(1,docuri.length);
       //alert("url: "+url+"  filename: "+filename+" slidedix: "+slideidx+" retain: "+retain);
       var msg = MLA.copyPasteSlideToActive(tmpPath, filename,slideidx, url, "oslo","oslo",retain);
}

function openPPTX(docuri)
{
       //alert("docuri for testOpen():"+ docuri);
       var tokens = docuri.split("/");
       var filename = tokens[tokens.length-1];
       var tmpPath = MLA.getTempPath(); 

       var config = MLA.getConfiguration();
       var fullurl= config.url;
       var url = fullurl + "/officesearch/download-support.xqy?uid="+docuri;
       var msg = MLA.openPPTX(tmpPath, docuri, url, "oslo","oslo");
}
/* -----------------------------HERE --------------------------------------------*/
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
