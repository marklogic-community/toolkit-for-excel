//Copyright 2002-2008 Mark Logic Corporation.  All Rights Reserved.
var MLA = {};

MLA.version = { "release" : "@MAJOR_VERSION.@MINOR_VERSION@PATCH_VERSION" }; 

MLA.getVersion = function()
{
	return MLA.version.release;
}

MLA.getCustomPieceIds = function()
{ 
	var pieces = window.external.getCustomPieceIds();

        var errStr = pieces.substring(0,6);
	var len = pieces.length;
        var errMsg = pieces.substring(7,len);

	if(errStr=="error:")
	  throw("Error: Not able to get CustomPieceIds; "+errMsg);

	var customPieces = pieces.split(" ");
	return customPieces;
}

MLA.getCustomPiece = function(customPieceId)
{
	var customPiece = window.external.getCustomPiece(customPieceId);

	var errStr = customPiece.substring(0,6);
	var len = customPiece.length;
        var errMsg = customPiece.substring(7,len);

	if(errStr=="error:")
	  throw("Error: Not able to getCustomPiece; "+errMsg);

	if(customPiece=="")
	  customPiece=null;
   
	return customPiece;
}

MLA.addCustomPiece = function(customPieceXml)
{
	var newId = window.external.addCustomPiece(customPieceXml);

	var errStr = newId.substring(0,6);
	var len = newId.length;
        var errMsg = newId.substring(7,len);

	if(errStr=="error:")
	  throw("Error: Not able to addCustomPiece; "+errMsg);

	if(newId =="")
	  newId=null;

	return newId;
}

MLA.deleteCustomPiece = function(customPieceId)
{
	var deletedPiece = window.external.deleteCustomPiece(customPieceId);

        var errStr = deletedPiece.substring(0,6);
	var len = deletedPiece.length;
        var errMsg = deletedPiece.substring(7,len);

        if(errStr=="error:")
          throw("Error: Not able to deleteCustomPiece; "+errMsg);
     
	if(deletedPiece=="")
	  deletedPiece = null;

        //return deletedPiece;
}



/*
MLA.getSelection = function()
{
	var selection = window.external.getSelection();

        var errStr = selection.substring(0,6);
	var len = selection.length;
        var errMsg = selection.substring(7,len);

	if(errStr == "error:")
   	   throw("Unable to getSelection: "+errMsg);

	var selections;
        if(selections == "")
	{
	   selections=null;
	}
	else
	{
	   selections = selection.split("U+016000");
	}

	return selections;
}
*/
MLA.getSelection = function()
{
	var arrCount=0;
	var selCount =0;
	var selections = new Array();
        
	var selection =  window.external.getSelection(selCount);

	var errStr = selection.substring(0,6);
	var len = 0;
        var errMsg = "";
	var err = false;

	if(errStr == "error:")
	{
		err=true;
		len = selection.length;
                errMsg = selection.substring(7,len);
		selection="";
	}


	selections[arrCount]=selection;

	while(selection!="")
	{
  	  selCount++;
          arrCount++;
	  selection = window.external.getSelection(selCount);

 	  var errStr = selection.substring(0,6);

	  if(errStr == "error:")
	  {
   	    err=true;
	    len = selection.length;
            errMsg = selection.substring(7,len);
	    selection="";

	  }

	  if(selection!="")
	      selections[arrCount] = selection;

	}

	if(err==true)
	{
	   throw("Error: Not able to getSelection; "+errMsg);
	}


	return selections;
}

MLA.getSentenceAtCursor = function()
{
	var rangeXml = window.external.getSentenceAtCursor();

        var errStr = rangeXml.substring(0,6);
	var len = rangeXml.length;
        var errMsg = rangeXml.substring(7,len);

        if(errStr == "error:")
	   throw("Error: Not able to get Sentence at Cursor; "+errMsg);

	return rangeXml;
}

MLA.getActiveDocStylesXml = function()
{ 
	var stylesXml = window.external.getActiveDocStylesXml();

	var errStr = stylesXml.substring(0,6);
	var len = stylesXml.length;
        var errMsg = stylesXml.substring(7,len);

	if(errStr=="error:")
          throw("Error: Not able to getActiveDocStylesXml; "+errMsg);

	if(stylesXml=="")
          stylesXml=null;

	return stylesXml;
}

MLA.isArray = function(obj)
{
 return obj.constructor == Array;
}

MLA.insertBlockContent = function(blockContentXml,stylesXml)
{
	if(stylesXml == null) 
	    stylesXml = "";
   
	var v_block="";
	var v_array = MLA.isArray(blockContentXml);
        //alert("ARRAY? :"+v_array);

	if(v_array)
	{
		for(var i=0;i<blockContentXml.length;i++)
		{
			v_block = v_block+blockContentXml[i];
		}
		//alert("v_block: "+v_block);
	}
	else
	{
		v_block = blockContentXml;
		//alert("v_block: "+v_block);
	}

        //var inserted = window.external.insertBlockContent(blockContentXml,stylesXml);
        var inserted = window.external.insertBlockContent(v_block,stylesXml);

	var errStr = inserted.substring(0,6);
	var len = inserted.length;
        var errMsg = inserted.substring(7,len);

	if(errStr=="error:")
	  throw("Error: Not able to insertBlockContent; "+errMsg);



	if(inserted=="")
	  inserted = null;

        //return inserted;
}

MLA.getConfiguration = function()
{
//	var configDetails = window.external.getConfiguration();

//	if(configDetails == "")
//	   throw(configDetails+":unable to getConfiguration.");

//	var configs = configDetails.split("U+016000");
        var version = window.external.getAddinVersion();
	var color = window.external.getOfficeColor();
	var webUrl = window.external.getBrowserUrl();

	if(version == "" || color == "" || webUrl == "")
		throw("Error: Not able to access configuration info.");

	MLA.config = {
		        "version":version,
			"url":webUrl,
			"theme":color
		       //"version":configs[0],
                       //"url":configs[1],
	               //"color":configs[2]
	};

        return MLA.config;	
}
