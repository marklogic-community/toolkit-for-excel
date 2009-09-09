window.onload=initPage;

var debug = false;

function initPage()
{
	if(debug)
	  alert("initPage() begin");

	var testResults1 =  insertJSONTable();
	var testResults2 =  insertImage();
	var testResults3 =  insertSlideRetain();
	var testResults4 =  insertSlideNoRetain();
	var testResults5 =  embedOLE();
        var testResults6 =  testAddCustomPiece();   
        var testResults7 =  testGetCustomPieceIds();   
	var testResults8 =  testGetCustomPiece(testResults7);
	var testResults9 =  convertFilenameToImageDir();
	var testResults10 = saveLocalCopy();
	var testResults11 = getTempPath();
	var testResults12 = getPresentationPath();
	var testResults13 = getPresentationName();
        var testResults14 = saveActivePresentation();
	var testResults15 = saveImages();
	var id = testGetCustomPieceIds()
	var testResults16 = testDeleteCustomPiece(id);
	var testResults17 = saveActivePresentationAndImages();
	   
	//can write to file, or just save one big xml file to ML
	var testOutput = generateTestTemplate(testResults1,testResults2,testResults3,testResults4,testResults5,testResults6,testResults7,testResults8,testResults9,testResults10,testResults11,testResults12,testResults13,testResults14,testResults15,testResults16,testResults17);
	writeToFile(testOutput);

	if(debug)
	  alert("initializing page");

}

function insertText()
{
	//alert("inserting text");
	var msg = MLA.insertText("I'm on a Boat!");
	return msg;

}

function insertJSONTable()
{
   //alert("in table function");

/*	var tbl =  "{\"defaultValue\":{"+
                    "\"columns\": { \"title\" : \"Title\","+ 
                                  "\"author\":\"Author\","+ 
                                  "\"isbn\": \"ISBN #\","+ 
                                  "\"description\":\"Description\"},"+
                     "\"rows\":["+
 "['JavaScript 101', 'Lu Sckrepter','4412', 'Some long description'],"+
 "['Ajax with Java', 'Jean Bean','4413', 'Some long description']"+
                            "]"+
                                 "}}";
				 */

	/*	var tbl =  "{\"defaultValue\":{"+
                    "\"columns\": ['Header1', 'Header2','Header3', 'Header4'],"+
                     "\"rows\":["+
 "['JavaScript 101', 'Lu Sckrepter','4412', 'Some long description'],"+
 "['Ajax with Java', 'Jean Bean','4413', 'Some long description']"+
                            "]"+
                                 "}}"; */
/*
		var tbl =  "{"+
                    "\"headers\": ['Header1', 'Header2','Header3', 'Header4']"+
                                 "}";
*/
	var tbl =  "{"+
                    "\"headers\": ['Header1', 'Header2','Header3', 'Header4'],"+
		     "\"values\": ["+
		                   "['JavaScript 101', 'Lu Sckrepter','4412', 'Some long description'],"+
		                   "['Ajax with Java', 'Jean Bean','4413', 'Some long description']"+
				   "]"+
                                 "}";
     if(debug)
	alert(tbl);
		
	var x = window.external.insertJSONTable(tbl);
	return x;
}

function useSaveFileDialog()
{
	var msg = MLA.useSaveFileDialog();

	if(debug)
	  alert("useSaveFileDialog() message: "+msg);
	
	return msg;

}

function convertFilenameToImageDir()
{
	var dirname = MLA.convertFilenameToImageDir("foo.pptx");

	if(debug)
	  alert("convertFIlenameToImageDir() filename: "+dirname);

	return dirname;

}

function getTempPath()
{
	var tmppath = MLA.getTempPath();

	if(debug)
          alert("getTempPath() tmppath: "+tmppath);
	
	return tmppath;
}

function getPresentationPath()
{
	var prespath = MLA.getPresentationPath();

	if(debug)
	  alert("getPresenationPath() prespath: "+prespath)
	 
	return prespath;
}

function getPresentationName()
{
	var pname = MLA.getPresentationName();

	if(debug)
	  alert("getPresentationName() pname: "+pname);

	return pname;
}

function saveLocalCopy()
{
        var tmppath = MLA.getTempPath();
	var filename = tmppath+"foobar.pptx";
	var cpy = MLA.saveLocalCopy(filename);

	if(debug)
	  alert("saveLocalCopy() cpy: "+ cpy);

	return cpy;
}

function saveActivePresentation()
{       var user = "oslo";
        var pwd = "oslo";
	var tmppath = MLA.getTempPath();
	var filename = tmppath+"foobar.pptx";
        var url = "http://localhost:8023/Samples/utils/upload.xqy?uid=/foobar.pptx";
 	var msg=MLA.saveActivePresentation(filename,url,user,pwd);

	if(debug)
	  alert("saveActivePresentation() msg: "+msg);

	return msg;
}

function saveImages()
{
	var user = "oslo";
        var pwd = "oslo";
	var tmppath = MLA.getTempPath();
	var imgdir = MLA.convertFilenameToImageDir("foobar.pptx");
	var fullimgdir=tmppath+imgdir;

	if(debug)
	  alert("full img dir"+fullimgdir);

        var url = "http://localhost:8023/Samples/utils/upload.xqy?uid=";

	var msg = MLA.saveImages(fullimgdir, url, user, pwd);

	if(debug)
          alert("saveImages() msg: "+msg);

	return msg;
}

function saveActivePresentationAndImages()
{
	var user = "oslo";
        var pwd = "oslo";
	var tmppath = MLA.getTempPath();
        var filename = "foobar2.pptx";
	var url = "http://localhost:8023/Samples/utils/upload.xqy?uid=";
	var msg=MLA.saveActivePresentationAndImages(tmppath, filename, url, user, pwd);

	if(debug)
          alert("saveActivePresentationAndImages() msg: "+msg);

	return msg;
}
//---------------------
//
function embedOLE()
{
	 var docname = "/musicCatalog.xlsx";
         var title = "musicCatalog.xlsx";
	 var tmpPath = MLA.getTempPath(); 
         var config = MLA.getConfiguration();
         var fullurl= config.url;
         var url = fullurl + "/officesearch/download-support.xqy?uid="+docname;
         //alert("tmppath: "+tmpPath+"\n  url: "+url+ "\n   title: "+title);
         var msg = MLA.embedOLE(tmpPath, title, url, "oslo","oslo");
	 return msg;
}

function insertImage()
{     
       //picuri
       var picuri = "/ackbar.jpg";

       var config = MLA.getConfiguration();
       var fullurl= config.url;
       var picuri = fullurl + "/search/download-support.xqy?uid="+picuri;
       var msg = MLA.insertImage(picuri,"oslo","oslo");
       return msg;
}

function insertSlideRetain()
{
       //docuri, slideidx, retainidx
       var docuri = "/one.pptx";
       var slideidx = "1";
       var retain = "true";

       var tmpPath = MLA.getTempPath();

       var config = MLA.getConfiguration();
       var fullurl= config.url;
       var url = fullurl + "/search/download-support.xqy?uid="+docuri;
      
       var tokens = docuri.split("/");
       var filename = tokens[tokens.length-1];

       if(debug)
         alert(filename); 
       
       var msg = MLA.insertSlide(tmpPath, filename,slideidx, url, "oslo","oslo",retain);
       return msg;
}

function insertSlideNoRetain()
{
      //docuri, slideidx, retainidx
       var docuri = "/testOne.pptx";
       var slideidx = "3";
       var retain = "false";

       var tmpPath = MLA.getTempPath();

       var config = MLA.getConfiguration();
       var fullurl= config.url;
       var url = fullurl + "/search/download-support.xqy?uid="+docuri;
      
       var tokens = docuri.split("/");
       var filename = tokens[tokens.length-1];

       if(debug)
         alert(filename); 

       var msg = MLA.insertSlide(tmpPath, filename,slideidx, url, "oslo","oslo",retain);
       return msg;
}
//---------------------
//---------------------
//---------------------

function testAddCustomPiece()
{
	try{
		var v_title="TITLE";
		var v_description="DESCRIPTION";
		var v_publisher="PUBLISHER";
		var v_identifier="IDENTIFIER";
	
		var customPiece = generateTemplate(v_title,v_description,v_publisher,v_identifier);
        	var newid = MLA.addCustomXMLPart(customPiece);
		//	alert("ID IS"+newid);
		return newid;
	}catch(err)
	{
		return "error: "+err.description;
	}
       
}

function testGetCustomPieceIds()
{
	try{
		var ids = MLA.getCustomXMLPartIds();
		return ids[0];
	}catch(err)
	{
		return "error: "+err.description;
	}
}

function testGetCustomPiece(cid)
{
	try{
		var piece = MLA.getCustomXMLPart(cid);
		//alert("PIECE"+piece.xml);
		return piece.xml;
	}catch(err)
	{
		return "error: "+err.description;
	}
}

function testDeleteCustomPiece(id)
{
	try{
		var deletedPiece = MLA.deleteCustomXMLPart(id);
		return deletedPiece;
	}catch(err)
	{
		return "error: "+err.description;
	}
}

function writeToFile(output)
{
  try {
   var fso = new ActiveXObject("Scripting.FileSystemObject");
   var a = fso.CreateTextFile("C:\\unitTestAddin\\outputs\\onload_testresults.txt", true);
   a.WriteLine(output);
   a.Close();
 }
 catch(err){
   var strErr = 'Error:';
   strErr += '\nNumber:' + err.number;
   strErr += '\nDescription:' + err.description;
   document.write(newid);
  }
}


function generateTemplate(title,description,publisher,id)
{
	 var v_template ="<metadata "+
           "xmlns='http://example.org/myapp/' "+
           "xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' "+
           "xsi:schemaLocation='http://example.org/myapp/ http://example.org/myapp/schema.xsd' "+
           "xmlns:dc='http://purl.org/dc/elements/1.1/'>"+
           "<dc:title>"+
             title+
           "</dc:title>"+
           "<dc:description>"+
	     description+
           "</dc:description>"+
           "<dc:publisher>"+
	     publisher+
           "</dc:publisher>"+
           "<dc:identifier>"+
             id+
           "</dc:identifier>"+
           "</metadata>";
	 return v_template;

}

function generateTestTemplate(test1Res,test2Res,test3Res, test4Res,test5Res, test6Res,test7Res,test8Res,test9Res,test10Res,test11Res,test12Res,test13Res,test14Res,test15Res,test16Res,test17Res, test18Res, test19Res)
{
  var v_testTemplate ="<tests t='http://marklogic.openxml.test'>"+
	                 "<test id='1'>"+test1Res+"</test>"+
	                 "<test id='2'>"+test2Res+"</test>"+
	                 "<test id='3'>"+test3Res+"</test>"+
	                 "<test id='4'>"+test4Res+"</test>"+
	                 "<test id='5'>"+test5Res+"</test>"+
	                 "<test id='6'>"+test6Res+"</test>"+
	                 "<test id='7'>"+test7Res+"</test>"+
	                 "<test id='8'>"+test8Res+"</test>"+
	                 "<test id='9'>"+test9Res+"</test>"+
	                 "<test id='10'>"+test10Res+"</test>"+
	                 "<test id='11'>"+test11Res+"</test>"+
	                 "<test id='12'>"+test12Res+"</test>"+
	                 "<test id='13'>"+test13Res+"</test>"+
	                 "<test id='14'>"+test14Res+"</test>"+
	                 "<test id='15'>"+test15Res+"</test>"+
	                 "<test id='16'>"+test16Res+"</test>"+
	                 "<test id='17'>"+test17Res+"</test>"+
	                 "<test id='18'>"+test18Res+"</test>"+
	                 "<test id='19'>"+test19Res+"</test>"+
	              "</tests>";
  return v_testTemplate;
	              
}

function updateMetadata(i)
{
	var edited = false;
   if(i==1)
   {
	if(debug)
           alert("Saving Custom Piece");
        
	var customPieceIds = MLA.getCustomPieceIds();
	var customPieceId = null;
	for(i=0;i<customPieceIds.length;i++)
	{
	  if(customPieceIds[i] == null || customPieceIds ==""){
		  //do nothing
	  }else{
	        customPieceId = customPieceIds[i];
		var delPiece = MLA.deleteCustomPiece(customPieceId);
		edited=true;
	  }
	        
	} 

	var v_title       = document.getElementById("ML-Title").value;
        var v_description = document.getElementById("ML-Desc").value;
        var v_publisher   = document.getElementById("ML-Publisher").value;
        var v_identifier  = document.getElementById("ML-Id").value;

	/*
	if(v_title=="" || v_title==null)
		v_title="Please Enter A Title";
	if(v_description=="" || v_description==null)
		v_description="Please Enter A Description";
	if(v_publisher=="" || v_publisher==null)
		v_publisher="Please Enter A Publisher";
	if(v_identifier=="" || v_identifier==null)
		v_identifier="Please Enter An Id";
        */

	var customPiece = generateTemplate(v_title,v_description,v_publisher,v_identifier);

	if(debug)
	   alert(customPiece);

        var newid = MLA.addCustomPiece(customPiece);

	if(edited){
 	 //alert("Metadata Edited"); 
         //added
	 document.getElementById("ML-Message").innerHTML = "Document Metadata Edited";
	}else{
	 document.getElementById("ML-Message").innerHTML = "Metadata Saved With Document";
	}
	 
		/*
	  alert("Existing Metadata in the Document was edited.");
	}else{
	  alert("Metadata Saved To Document.");
	}*/   
   }
   else
   {    if(debug)
	   alert("Removing Custom Piece");
	var customPieceIds = MLA.getCustomPieceIds();
	var customPieceId = null;
	for(i=0;i<customPieceIds.length;i++)
	{
	  if(customPieceIds[i] == null || customPieceIds ==""){
		  //do nothing
	  }else{
	        customPieceId = customPieceIds[i];
		var delPiece = MLA.deleteCustomPiece(customPieceId);
	  }
	        
	}

       	document.getElementById("ML-Title").value="";
        document.getElementById("ML-Desc").value="";
        document.getElementById("ML-Publisher").value="";
        document.getElementById("ML-Id").value="";	
        document.getElementById("ML-Message").innerHTML = "No Metadata Saved with Document";
   }
}
