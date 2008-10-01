window.onload=initPage;

var debug = true;

function initPage()
{

	if(debug)
	  alert("initializing page");

	var customPieceIds = MLA.getCustomPieceIds();
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
		alert("IN IF");
            var xmlDoc = tmpCustomPieceXml;
            // xmlDoc=new ActiveXObject("Microsoft.XMLDOM");
            // xmlDoc.async="false";
            // xmlDoc.loadXML(tmpCustomPieceXml);
             var v_title       = xmlDoc.getElementsByTagName("dc:title")[0].childNodes[0].nodeValue;
             var v_description = xmlDoc.getElementsByTagName("dc:description")[0].childNodes[0].nodeValue;
             var v_publisher   = xmlDoc.getElementsByTagName("dc:publisher")[0].childNodes[0].nodeValue;
             var v_identifier  = xmlDoc.getElementsByTagName("dc:identifier")[0].childNodes[0].nodeValue;
             document.getElementById("v_title").value = v_title;
             document.getElementById("v_desc").value  = v_description;
             document.getElementById("v_pub").value   = v_publisher;
             document.getElementById("v_id").value    = v_identifier;
	     document.getElementById("v_fc").innerHTML = "Metadata Saved with Document";

	}else
	{ 
              document.getElementById("v_fc").innerHTML="No Metadata Saved with Document";
	//	alert("TEST" +x);
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

	var v_title       = document.getElementById("v_title").value;
        var v_description = document.getElementById("v_desc").value;
        var v_publisher   = document.getElementById("v_pub").value;
        var v_identifier  = document.getElementById("v_id").value;

	if(v_title=="" || v_title==null)
		v_title="Please Enter A Title";
	if(v_description=="" || v_description==null)
		v_description="Please Enter A Description";
	if(v_publisher=="" || v_publisher==null)
		v_publisher="Please Enter A Publisher";
	if(v_identifier=="" || v_identifier==null)
		v_identifier="Please Enter An Id";

	var customPiece = generateTemplate(v_title,v_description,v_publisher,v_identifier);

	if(debug)
	   alert(customPiece);

        var newid = MLA.addCustomPiece(customPiece);

	if(edited){
		alert("Metadata Edited"); 
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

   }
}
