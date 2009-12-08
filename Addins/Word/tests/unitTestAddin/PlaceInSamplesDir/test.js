window.onload=initPage;

var debug = false;

function initPage()
{
        var test1Results = testInsertText();
	var test2Results = testAddCustomPiece();
	var test3Results = testGetCustomPieceIds();
	var test4Results = testGetCustomPiece(test3Results);
	var test5Results = testDeleteCustomPiece(test3Results);
	var test6Results = testCreateParagraph("THIS IS A TEST");
	var test7Results = testInsertBlockContent(test6Results);

	//can write to file, or just save one big xml file to ML
	var testOutput = generateTestTemplate(test1Results,test2Results,test3Results,test4Results.xml,test5Results,test6Results.xml, test7Results);
	writeToFile(testOutput);

	if(debug)
	  alert("initializing page");

/*	var customPieceIds = MLA.getCustomPieceIds();
	var customPieceId = null;
	var tmpCustomPieceXml = null;
	for(i=0;i<customPieceIds.length;i++)
	{
	  if(customPieceIds[i] == null || customPieceIds ==""){
	     // do nothing
	  }else{

		if(debug)
		   alert("PIECE ID: "+customPieceIds[i]);

	        customPieceId = customPieceIds[i];
		tmpCustomPieceXml = MLA.getCustomPiece(customPieceId);
		if(debug)
		   alert(tmpCustomPieceXml.xml);
	  }
	        
	}

	if(tmpCustomPieceXml != null)// && tmpCustomPieceXml.length > 1)
	{
	    //alert("IN IF");
            var xmlDoc = tmpCustomPieceXml;
            // xmlDoc=new ActiveXObject("Microsoft.XMLDOM");
            // xmlDoc.async="false";
            // xmlDoc.loadXML(tmpCustomPieceXml);
	         	var v_title="";
			var v_description="";
			var v_publisher="";
			var v_identifier="";

			if(xmlDoc.getElementsByTagName("dc:title")[0].hasChildNodes()) 
			   v_title = xmlDoc.getElementsByTagName("dc:title")[0].childNodes[0].nodeValue;

			if(xmlDoc.getElementsByTagName("dc:description")[0].hasChildNodes()) 
			   v_description = xmlDoc.getElementsByTagName("dc:description")[0].childNodes[0].nodeValue;

			if(xmlDoc.getElementsByTagName("dc:publisher")[0].hasChildNodes()) 
			   v_publisher = xmlDoc.getElementsByTagName("dc:publisher")[0].childNodes[0].nodeValue;
			
			if(xmlDoc.getElementsByTagName("dc:identifier")[0].hasChildNodes()) 
			   v_identifier = xmlDoc.getElementsByTagName("dc:identifier")[0].childNodes[0].nodeValue;
*/
          /*   var v_title       = xmlDoc.getElementsByTagName("dc:title")[0].childNodes[0].nodeValue;
             var v_description = xmlDoc.getElementsByTagName("dc:description")[0].childNodes[0].nodeValue;
             var v_publisher   = xmlDoc.getElementsByTagName("dc:publisher")[0].childNodes[0].nodeValue;
             var v_identifier  = xmlDoc.getElementsByTagName("dc:identifier")[0].childNodes[0].nodeValue;
	  */
/*             
             document.getElementById("ML-Title").value = v_title;
             document.getElementById("ML-Desc").value  = v_description;
             document.getElementById("ML-Publisher").value   = v_publisher;
             document.getElementById("ML-Id").value    = v_identifier;
	    
	     document.getElementById("ML-Message").innerHTML = "Metadata Saved with Document";
	     

	}else
	{ 
              document.getElementById("ML-Message").innerHTML="No Metadata Saved with Document";
	//	alert("TEST");
	}
*/
}

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

function writeToFile(output)
{
  try {
   var fso = new ActiveXObject("Scripting.FileSystemObject");
   var a = fso.CreateTextFile("C:\\testfile.txt", true);
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
