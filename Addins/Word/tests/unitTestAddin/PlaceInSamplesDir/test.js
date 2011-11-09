window.onload=initPage;

var debug = false;

function initPage()
{
	var testOutput="";
	var fName ="";
        //alert(MLA.getDocumentName());
	var docname = MLA.getDocumentName();
        //alert("DOCNAME: "+docname);
	if(docname == "maptest.docx") //ccMapTest.docx
	{
	   //var tag = "pttesttag";
           //var tag = "3495257";
           var tag = "19232534";
	   var ids = MLA.getCustomXMLPartIds();
           var xpath = "dc:metadata/dc:title";
           var prefix = "xmlns:dc='http://purl.org/dc/elements/1.1/'";
	   //alert("HERE");
	   testOutput = MLA.mapContentControl(tag, xpath, prefix, ids[0])
	   fName = "maptest.txt";
	}
	else if(docname == "addcontrolstest.docx")
	{
	  var test1Results = addControlWithImage(); 
          var test2Results = addControlWithText();
	  var test3Results = addControlWithList();

	  testOutput = "<tests>"+
		         "<test>"+test1Results+"</test>"+
		         "<test>"+test2Results+"</test>"+
		         "<test>"+test3Results+"</test>"+
		       "</tests>";
	  //alert("test output"+testOutput);

	  fName="addcontrolstest.txt";
	}
	else if(docname == "lockcontroltest.docx")
	{

	  lockControl();
	  unlockControlContents();
	  MLA.setContentControlFocus("21730512"); 
	  MLA.insertText("This is another test.");
	  lockControlContents();

	  var test1Results = "test completed";

	  testOutput = "<tests>"+
		         "<test>"+test1Results+"</test>"+
		       "</tests>";
	  fName = "lockcontroltest.txt";
	}
	else if(docname == "controlstest.docx")
	{
	  var test1Results = MLA.getTempPath();
	  var test2Results = MLA.getDocumentPath();
	  var test3Results = MLA.getDocumentName();
          var test4Results = MLA.addContentControl("FOOBAR","FANCYTITLE","wdContentControlRichText","false","");
	  var test5Results = addComplexControl();

	  MLA.setContentControlFocus(test5Results);
	  var info = MLA.getParentContentControlInfo();
	  var test6Results = "ID: "+info.id + "  tag: " +info.tag+"  title: "+info.title+
		             " type: "+info.type+" parentTag: "+info.parentTag+
			     " parentID: "+info.parentID;

	  MLA.setContentControlFocus(test4Results);
	  var test7Results = insertXML();
	  //var test8Results = listContentControls();

	  testOutput = "<tests>"+
		         "<test>"+test1Results+"</test>"+
		         "<test>"+test2Results+"</test>"+
		         "<test>"+test3Results+"</test>"+
		         "<test>"+test4Results+"</test>"+
		         "<test>"+test5Results+"</test>"+
		         "<test>"+test6Results+"</test>"+
		         "<test>"+test7Results+"</test>"+
		       "</tests>";

	  fName = "controlstest.txt";
	  
	  
	}
	else if(docname == "gettexttest.docx")
	{
	  displayControlRange();
	  var test1Results = setControlStyle();
	  var test2Results = setControlTag();
	  var test3Results = setControlTitle();
	  setControlFocus();
	  var test4Results = getControlText();
	  var test5Results = getControlXML();
	  hideControlRange();

	  testOutput = "<tests>"+
		         "<test>"+test1Results+"</test>"+
		         "<test>"+test2Results+"</test>"+
		         "<test>"+test3Results+"</test>"+
		         "<test>"+test4Results+"</test>"+
		         "<test>"+test5Results.text+"</test>"+
		       "</tests>";

	  fName="gettexttest.txt";
	}
	else
	{
          var test1Results = testInsertText();
	  var test2Results = testAddCustomPiece();
	  var test3Results = testGetCustomPieceIds();
	  var test4Results = testGetCustomPiece(test3Results);
	  var test5Results = testDeleteCustomPiece(test3Results);
	  var test6Results = testCreateParagraph("THIS IS A TEST");
	  var test7Results = testInsertBlockContent(test6Results);

	//can write to file, or just save one big xml file to ML
	  testOutput = generateTestTemplate(test1Results,test2Results,test3Results,test4Results.xml,test5Results,test6Results.xml, test7Results);
	  fName="originaltests.txt";
	}

	writeToFile(testOutput,fName);

	if(debug)
	  alert("initializing page");

}

/* ================ BEGIN ADD CONTROLS TESTS ================================*/
function addControlWithImage()
{
       var pic = "/ackbar.jpg";
       var config = MLA.getConfiguration();
       var fullurl= config.url;
       var picuri = fullurl + "/download-support.xqy?uid="+pic;

       var msg2 = MLA.addContentControl("FOOBAR2","FANCYTITLE2","wdContentControlPicture","true");
       //var msg = MLA.insertContentControlImage(msg2, picuri,"oslo","oslo");
       return msg2;

}

function addControlWithText()
{
	var msg8=MLA.addContentControl("BAR8","FANCYTITLE7","wdContentControlRichText","false");
        MLA.insertContentControlText(msg8,"FOOBAR");
	return msg8;
}

function addControlWithList()
{
        var dropdownid = MLA.addContentControl("FOOBAR3","FANCYTITLE3","wdContentControlDropdownList","true");
        var msg =  MLA.addContentControlDropDownListEntries(dropdownid,"FOO","BAR",0);
	var msg2 = MLA.addContentControlDropDownListEntries(dropdownid,"FOO2","BAR2",0);
	var msg3 = MLA.addContentControlDropDownListEntries(dropdownid,"FOO3","BAR3",0);
	var msg4 = MLA.addContentControlDropDownListEntries(dropdownid,"FOO4","BAR4",0);
	var msg5 = MLA.addContentControlDropDownListEntries(dropdownid,"FOO6","BAR56",0);
	return dropdownid;
}

/* ================ END ADD CONTROLS TESTS ==================================*/

/* ================ BEGIN LOCK CONTROLS TESTS ===============================*/
function lockControl()
{
	var msg = MLA.lockContentControl("21730501");
	var msg = MLA.lockContentControl("21730512");
}

function unlockControl()
{
	var msg = MLA.unlockContentControl("21730501");
	var msg = MLA.unlockContentControl("21730512");
}

function lockControlContents()
{
	var msg = MLA.lockContentControlContents("21730501");
	var msg2 = MLA.lockContentControlContents("21730512");
}

function unlockControlContents()
{
	var msg =  MLA.unlockContentControlContents("21730501");
	var msg2 = MLA.unlockContentControlContents("21730512");
}

/* ================ END LOCK CONTROLS TESTS =================================*/

/* ================ BEGIN CONTROLS TESTS ====================================*/
function addComplexControl()
{       //alert("HERE");
	var msg15=MLA.addContentControl("FOOBAR","FANCYTITLE","wdContentControlRichText","true","");
	var parentID = msg15;
	var msg16=MLA.addContentControl("BAR16","FANCYTITLE6","wdContentControlRichText","true",parentID);
        var secondParentID=msg16;
	var msg17=MLA.addContentControl("BAR17","FANCYTITLE7","wdContentControlRichText","true",parentID);
	

	var msgg1 = MLA.setContentControlPlaceholderText(msg16,"THIS IS MY MESSAGE 1","true");
	var msgg2 = MLA.setContentControlPlaceholderText(msg17,"THIS IS MY MESSAGE 2");

	var msg18=MLA.addContentControl("FOO16A","FANCY16A","wdContentControlRichText","false",secondParentID);
	var msg19=MLA.addContentControl("FOO16B","FANCY16B","wdContentControlRichText","false",secondParentID);

	//alert(msg19);
	return msg19;

}

function insertXML()
{
	url  =  "http://localhost:8023/wordTests/fetchWordOpenXml.xqy";
        var opc_xml = loadXMLDoc(url);
        //alert(opc_xml);
	MLA.insertWordOpenXML(opc_xml);

	var mydom = MLA.createXMLDOM(opc_xml);
	MLA.insertWordOpenXML(mydom);
	return "inserted WordOpenXML both text and DOM style";

}

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

function processPostReqChange() 
{
    if (req2.readyState == 4) {
        if (req2.status == 200) {
            response = req2.responseText;
        } else {
            alert("There was a problem retrieving the XML data:\n" + req2.statusText);
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
/* ================ END CONTROLS TESTS ======================================*/

/* ================ BEGIN GETTEXT TESTS =====================================*/
function setControlStyle()
{
	//MLA.setContentControlStyle("13255863", "test");
	MLA.setContentControlStyle("13255863", "Heading 1");
	MLA.setContentControlStyle("13255870", "Subtitle");
	return "set style";
}

function setControlTag()
{
	MLA.setContentControlTag("13255863", "MYAWESOMETAG");
	return "set tag";
}

function setControlTitle()
{
	MLA.setContentControlTitle("13255863", "MYAWESOMETITLE");
	return "set title";
}

function setControlFocus()
{
	var msg = MLA.setContentControlFocus("13255863");
	//window.external.setContentControlFocus("BAR6");
	return "set focus";
}

function getControlText()
{
        var txt = MLA.getContentControlText("13255863");
	//alert(txt);
	return txt;
}

function getControlXML()
{
	var myxml =MLA.getContentControlWordOpenXML("13255863");
	//alert(myxml.xml);
        return myxml;
}

function hideControlRange()
{
	var hidden = MLA.hideContentControlRange("13255863");
	return hidden;
}

function displayControlRange()
{
	
	var displayed = MLA.displayContentControlRange("13255863");
	return displayed;
}


/* ================ END GETTEXT TESTS =======================================*/

/* ================ BEGIN ORIGINAL TESTS ====================================*/

function testInsertText()
{
	var result = MLA.insertText("TEST");
	return result;
}

function testAddCustomPiece()
{
	var v_title="TITLE";
	var v_description="DESCRIPTION";
	var v_publisher="PUBLISHER";
	var v_identifier="IDENTIFIER";
        document.getElementById("ML-Title").value = v_title;
        document.getElementById("ML-Desc").value  = v_description;
        document.getElementById("ML-Publisher").value   = v_publisher;
        document.getElementById("ML-Id").value    = v_identifier;
	
	var customPiece = generateTemplate(v_title,v_description,v_publisher,v_identifier);
        var newid = MLA.addCustomXMLPart(customPiece);
	return newid;
       
}

function testGetCustomPieceIds()
{
	var ids = MLA.getCustomXMLPartIds();
	return ids[0];
}

function testGetCustomPiece(cid)
{
	var piece = MLA.getCustomXMLPart(cid);
	return piece;
}

function testDeleteCustomPiece(id)
{
	var deletedPiece = MLA.deleteCustomXMLPart(id);
	return deletedPiece;
}

function testCreateParagraph(text)
{
	var para = MLA.createParagraph(text);
	return para;
}

function testInsertBlockContent(block)
{
	var bret = MLA.insertBlockContent(block);
	return bret;
}

/* ================ END ORIGINAL TESTS ======================================*/

function writeToFile(output, filename)
{
  try {
   var fso = new ActiveXObject("Scripting.FileSystemObject");
   var a = fso.CreateTextFile("C:\\Users\\paven\\AppData\\Local\\Temp\\junk\\"+filename, true);
   a.WriteLine(output);
   a.Close();
 }
 catch(err){
   var strErr = 'Error:';
   strErr += '\nNumber:' + err.number;
   strErr += '\nDescription:' + err.description;
   alert("ERROR: "+strErr);
   document.write(newid);
  }
}


function generateTemplate(title,description,publisher,id)
{
	 var v_template ="<dc:metadata "+
           //"xmlns='http://example.org/myapp/' "+
           //"xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' "+
           //"xsi:schemaLocation='http://example.org/myapp/ http://example.org/myapp/schema.xsd' "+
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
           "</dc:metadata>";
	 return v_template;

}

function generateTestTemplate(test1Res,test2Res,test3Res, test4Res,test5Res, test6Res,test7Res)
{
  var v_testTemplate ="<tests t='http://marklogic.openxml.test'>"+
	                 "<test id='1'>"+test1Res+"</test>"+
	                 "<test id='2'>"+test2Res+"</test>"+
	                 "<test id='3'>"+test3Res+"</test>"+
	                 "<test id='4'>"+test4Res+"</test>"+
	                 "<test id='5'>"+test5Res+"</test>"+
	                 "<test id='6'>"+test6Res+"</test>"+
	                 "<test id='7'>"+test7Res+"</test>"+
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
        
	var customPieceIds = MLA.getCustomXMLPartIds();
	var customPieceId = null;
	for(i=0;i<customPieceIds.length;i++)
	{
	  if(customPieceIds[i] == null || customPieceIds ==""){
		  //do nothing
	  }else{
	        customPieceId = customPieceIds[i];
		var delPiece = MLA.deleteCustomXMLPart(customPieceId);
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

        var newid = MLA.addCustomXMLPart(customPiece);

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
	var customPieceIds = MLA.getCustomXMLPartIds();
	var customPieceId = null;
	for(i=0;i<customPieceIds.length;i++)
	{
	  if(customPieceIds[i] == null || customPieceIds ==""){
		  //do nothing
	  }else{
	        customPieceId = customPieceIds[i];
		var delPiece = MLA.deleteCustomXMLPart(customPieceId);
	  }
	        
	}

       	document.getElementById("ML-Title").value="";
        document.getElementById("ML-Desc").value="";
        document.getElementById("ML-Publisher").value="";
        document.getElementById("ML-Id").value="";	
        document.getElementById("ML-Message").innerHTML = "No Metadata Saved with Document";
   }
}
