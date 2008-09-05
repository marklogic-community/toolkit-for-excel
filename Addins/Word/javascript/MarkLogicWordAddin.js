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
	var customPieces = pieces.split("U+016000");

	if(customPieces == "") 
	  customPieces =  null;

	if(customPieces=="error")
	  throw(customPieces+": unable to get CustomPieceIds.");

	//for(var i=0;i<customPieces.length;i++)
	//	alert("ID: "+customPieces[i]+"TEST");
	
	return customPieces;
}

MLA.getCustomPiece = function(customPieceId)
{
	var customPiece = window.external.getCustomPiece(customPieceId);

	if(customPiece=="")
	  customPiece=null;
   
	if(customPiece=="error")
	  throw(customPiece+": unable to getCustomPiece.");

	return customPiece;
}

MLA.addCustomPiece = function(customPieceXml)
{
	var newId = window.external.addCustomPiece(customPieceXml);

	if(newId =="")
	  newId=null;

	if(newId=="error")
	  throw(newId+": unable to addCustomPiece.");

	return newId;
}

MLA.deleteCustomPiece = function(customPieceId)
{
	var deletedPiece = window.external.deleteCustomPiece(customPieceId);
     
	if(deletedPiece=="")
	  deletedPiece = null;

	if(deletedPiece=="error")
          throw(deletedPiece+": unable to deleteCustomPiece.");

        //return deletedPiece;
}



MLA.getSelection = function()
{
	var selection = window.external.getSelection();
	var selections;
        if(selections == "")
	{
	   selections=null;
	}
	else
	{
	   selections = selection.split("U+016000");
	}


	if(selections[0] == "error")
   	   throw(selections[0]+": unable to getSelection.");

	return selections;
}

MLA.getRangePreview = function()
{
	var rangeXml = window.external.getRangePreview();
   
	if(rangeXml == "error")
	   throw(rangeXml+": unable to getCursorXml.");

	return rangeXml;

}

MLA.getActiveDocStylesXml = function()
{ 
	var stylesXml = window.external.getActiveDocStylesXml();

	if(stylesXml=="")
          stylesXml=null;

	if(stylesXml=="error")
          throw(stylesXml+": unable to getActiveDocStylesXml.");

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

	if(inserted=="")
	  inserted = null;
	
	if(inserted=="error")
	  throw(inserted+": unable to insertBlockContent.");

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
		throw("unable to access configuration");

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
