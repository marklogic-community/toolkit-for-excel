//Copyright 2002-2008 Mark Logic Corporation.  All Rights Reserved.
var MLA = {};

MLA.version = { "release" : "@MAJOR_VERSION.@MINOR_VERSION@PATCH_VERSION" }; 

MLA.getVersion = function()
{
	return MLA.version.release;
}

MLA.errorCheck = function(message)
{
	var returnVal = null;
        var errStr = message.substring(0,6);
	var len = message.length;
        var errMsg = message.substring(7,len);

        if(errStr=="error:")
		returnVal=errMsg;

	return returnVal;

}

MLA.getCustomPieceIds = function()
{ 
	var pieces = window.external.getCustomPieceIds();

	var errMsg = MLA.errorCheck(pieces);
	if(errMsg!=null)
	  throw("Error: Not able to get CustomPieceIds; "+errMsg);

	var customPieces = pieces.split(" ");
	return customPieces;
}

MLA.getCustomPiece = function(customPieceId)
{
	var customPiece = window.external.getCustomPiece(customPieceId);

	var errMsg = MLA.errorCheck(customPiece);
	if(errMsg!=null)
	   throw("Error: Not able to getCustomPiece; "+errMsg);

	if(customPiece=="")
	  customPiece=null;
   
	return customPiece;
}

MLA.addCustomPiece = function(customPieceXml)
{
	var newId = window.external.addCustomPiece(customPieceXml);

	var errMsg = MLA.errorCheck(newId);
	if(errMsg!=null)
	   throw("Error: Not able to addCustomPiece; "+errMsg);

	if(newId =="")
	  newId=null;

	return newId;
}

MLA.deleteCustomPiece = function(customPieceId)
{
	var deletedPiece = window.external.deleteCustomPiece(customPieceId);

        var errMsg = MLA.errorCheck(deletedPiece);
	if(errMsg!=null)
	   throw("Error: Not able to deleteCustomPiece; "+errMsg);
     
	if(deletedPiece=="")
	  deletedPiece = null;

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

	var err = false;
	var errMsg = MLA.errorCheck(selection);
	if(errMsg!=null)
	{
		err=true;
		selection="";
	}


	selections[arrCount]=selection;

	while(selection!="")
	{
  	  selCount++;
          arrCount++;
	  selection = window.external.getSelection(selCount);


	  var errMsg = MLA.errorCheck(selection);
	  if(errMsg!=null){
   	    err=true;
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

	var errMsg = MLA.errorCheck(rangeXml);
	if(errMsg!=null) 
	   throw("Error: Not able to get Sentence at Cursor; "+errMsg);

	return rangeXml;
}

MLA.getActiveDocStylesXml = function()
{ 
	var stylesXml = window.external.getActiveDocStylesXml();

        var errMsg = MLA.errorCheck(stylesXml);
	if(errMsg!=null)
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

	if(v_array)
	{
		for(var i=0;i<blockContentXml.length;i++)
		{
			v_block = v_block+blockContentXml[i];
		}
	}
	else
	{
		v_block = blockContentXml;
	}

        var inserted = window.external.insertBlockContent(v_block,stylesXml);

	var errMsg = MLA.errorCheck(inserted);
	if(errMsg!=null)
	   throw("Error: Not able to insertBlockContent; "+errMsg);

	if(inserted=="")
	  inserted = null;

}

MLA.getConfiguration = function()
{
        var version = window.external.getAddinVersion();
	var color = window.external.getOfficeColor();
	var webUrl = window.external.getBrowserURL();

	if(version == "" || color == "" || webUrl == "")
		throw("Error: Not able to access configuration info.");

	MLA.config = {
		        "version":version,
			"url":webUrl,
			"theme":color
	};

        return MLA.config;	
}

MLA.addCommentForText = function(searchText, commentText)
{
	var commentAdded = window.external.addCommentForText(searchText, commentText);

        var errMsg = MLA.errorCheck(commentAdded);
	if(errMsg!=null)
	   throw("Error: Not able to add Comment for text "+errMsg);
     
	if(commentAdded=="")
	  commentAdded = null;
}

MLA.insertText = function(textToInsert)
{
	var textAdded = window.external.insertText(textToInsert);
	var errMsg = MLA.errorCheck(textAdded);
	if(errMsg!=null)
	   throw("Error: Not able to insert text "+errMsg);
     
	if(textAdded=="")
	  textAdded = null;
}
