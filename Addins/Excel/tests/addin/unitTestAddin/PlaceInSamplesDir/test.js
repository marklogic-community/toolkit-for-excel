window.onload=initPage;

var debug = false;

function initPage()
{
	//alert("TEST");
        var test1Results = testGetWBName(); //workbook name
	//alert("WB NAME IN ADDIN:  " +test1Results);

	var test2Results = testGetWSName(); //worksheet name
	var test3Results = testGetWBNames(); //workbook names (add wb first)
	var test4Results = testGetWBWSNames(); //workbook worksheet names
	var test5Results = "";// testAddWorkbook(); //add workbook
	var test6Results = "";//testSetActiveWB(); //set active wb (may have to remove, or take care of in .sln that launches excel? )
	var test7Results = testAddWorksheet(); //add worksheet
        var test8Results = testSetActiveWS(); // set active worksheet
	var test9Results =  testCells();       //populate worksheet
        var test10Results = testAddNamedRange();  //add 2 named ranges, one for everything, one for small range
        var test11Results = testAddAutoFilter();  //add autofilter to all
        var test12Results = testGetRangeNames();  //get range names
        var test13Results = testSetRangeByName();  //set range to smaller
	var test14Results = testClearNamedRange(); //clear range we just set too
        var test15Results = testClearRangeTest();  //clear random range 
        var test16Results = testRemoveRanges();    //remove smaller named range
        var test17Results = testAddCustomPiece();    //remove smaller named range
        var test18Results = testGetCustomPieceIds();    //remove smaller named range
	var test19Results = testGetCustomPiece(test18Results);
// testClearWorksheet();
//testSelectedRange();
//testSelectedCells();
	   
	//can write to file, or just save one big xml file to ML
	var testOutput = generateTestTemplate(test1Results,test2Results,test3Results,test4Results,test5Results,test6Results, test7Results, test8Results, test9Results,test10Results,test11Results,test12Results,test13Results,test14Results,test15Results,test16Results,test17Results,test18Results,test19Results);
	writeToFile(testOutput);

	if(debug)
	  alert("initializing page");

}

function testGetWBName()
{
	try{
             	return MLA.getActiveWorkbookName();
	}catch(err)
	{
	     	return "error: "+err.description;
	}
}

function testGetWSName()
{
	try{
             	return MLA.getActiveWorksheetName();
	}catch(err)
	{
	     	return "error: "+err.description;
	}
}
 
function  testGetWBNames()
{

	try{
    	     	var names = MLA.getAllWorkbookNames();
    	     	var wb  = "";

             	for(var i=0; i < names.length; i++) 
	        	wb = wb+names[i]+",";

             	return wb+"end";
	}catch(err)
  	{
    	     	return "error: "+err.description;
  	}	

}

function testGetWBWSNames()
{
	try{
             	var names = MLA.getActiveWorkbookWorksheetNames();
             	var wss = "";

             	for(var i=0; i < names.length; i++){
	        	wss = wss+names[i]+",";
	   //alert("WSS: "+ wss);
             }

	     	return wss+"end";

	}catch(err)
	{
	     	return "error: "+err.description;
	}
     
}

function testAddWorkbook()
{
	try{
             	var x = MLA.addWorkbook("FOO","BAR","FUBAR");
             	return x;
	}catch(err)
	{
		return "error: "+err.description;
	}
}

function testAddWorksheet()
{
	try{
                var x ="FUBAR";
                var msg=MLA.addWorksheet(x);
                return msg;
	}catch(err)
	{
		return "error: "+err.description;
	}
}

function testSetActiveWB()
{
	try{
    	        var test = MLA.setActiveWorkbook("test");  //Book1
    		return "added workbook";
	}catch(err)
	{
		return "error: "+err.description;
	}
}

function testSetActiveWS()
{
	try{
    		var test = MLA.setActiveWorksheet("Sheet1");  //Sheet3
    		return "added worksheet";
	}catch(err)
	{
		return "error: "+err.description;
	}
    //alert("message is:"+test)
}

function testCells()
{
	try{
     		var cells = new Array();
		//alert("array length"+cells.length);
     		for(c = 1; c < 10; c++)
     		{ 
			for(r = 1; r < 20; r++)
			{
          			var cell = new MLA.Cell(r,c);
	  			cell.value2 = 999;
	  			cells.push(cell);
			}
     		}
		//alert("FINAL array length: "+cells.length);
     		var v_msg = MLA.setCellValue(cells);
     		return v_msg;
	}catch(err)
	{
		return "error: "+err.description;
	}
}

function testAddNamedRange()
{
	try{
     		var nmdRange1 = MLA.addNamedRange("$A$1","$I$19","Everything");
     		var nmdRange2 = MLA.addNamedRange("$A$1","$A$10","MyRange");
     		return nmdRange1;
	}catch(err)
	{
		return "error: "+err.description;
	}
}

function testAddAutoFilter()
{
	try{
     		var filter = MLA.addAutoFilter("$A1", "$I19");
     		return filter;
	}catch(err)
	{
		return "error: "+err.description;
	}
}

function testGetRangeNames()
{
	try{
		//alert("GETTING NAMES");
		var names = MLA.getNamedRangeNames();
		var ranges="";
		for(i=0;i<names.length;i++)
			ranges =ranges+names[i]+",";

		return ranges+"end";
	}catch(err)
	{
		return "error: "+err.description;
	}
}

function testSetRangeByName()
{
	try{

		var msg = MLA.setActiveRangeByName("MyRange");
		return msg;
	}catch(err)
	{
		return "error: "+err.description;
	}
}

function testClearNamedRange()
{
	try{
         	var x = MLA.clearNamedRange("MyRange");
	 	return x;
	}catch(err)
	{
		return "error: "+err.description;
	}
}

function testClearRangeTest()
{
	try{
		var msg=MLA.clearRange("A1","C3");
		return msg;
	}catch(err)
	{
		return "error: "+err.description;
	}
}

function testRemoveRanges()
{
	try{
		var msg = MLA.removeNamedRange("MyRange");
        	return msg;
	}catch(err)
	{
		return "error: "+err.description;
	}
}

// testClearWorksheet();
// testSelectedRange();
// testSelectedCells();

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
   var a = fso.CreateTextFile("C:\\unitTestAddin\\outputs\\testresults.txt", true);
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
