/* 
Copyright 2009 Mark Logic Corporation

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
*/

/** 
 * @fileoverview  API documentation for MarkLogicPowerPointAddin.js
 *
 *
 *
 * {@link http://www.marklogic.com} 
 *
 * @author Pete Aven pete.aven@marklogic.com
 * @version 0.1 
 */


/**
 * The MLA namespace is used for global attribution. The methods within this namespace provide ways of interacting with an active Open XML document through a WebBrowser control. The control must be deployed within an Addin in Office 2007.
 *
 * The functions here provide ways for interacting with the active presentation in PowerPoint ; however, the functions getCustomXMLPart(), getCustomXMLPartIds(), addCustomXMLPart(), and deleteCustomXMLPart() will work for any Open XML package, provided they are used within the context of an Addin for the appropriate Office application.
 */
var MLA = {};
/*
function MLA(){
      this.getClassName = function(){
      return "MLA";
   
      }
}
*/
/** @ignore */
MLA.version = { "release" : "@MAJOR_VERSION.@MINOR_VERSION@PATCH_VERSION" }; 

/** @ignore */
MLA.SimpleRange = function(begin,finish){
	this.start = begin;
	this.end = finish;

};

/** @ignore */
String.prototype.trim = function() {
	return this.replace(/^\s+|\s+$/g,"");
}


/**
 * Returns version of MarkLogicPowerPointAddin.js library.
 * @return the version of MarkLogicWordAddin.js
 * @type String
 */
MLA.getVersion = function()
{
	return MLA.version.release;
}

/** @ignore */
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
/** Utility function for creating Microsoft.XMLDOM object from string.
 *
 *@param xmlString the string to be loaded into a XMLDOM object.  The string must be serialized, well-formed XML.
 *@return Microsoft.XMLDOM object
 *@throws Exception if unable to create the XMLDOM object
 */
MLA.createXMLDOM = function(xmlstring)
{
	try{
		var xmlDom = new ActiveXObject("Microsoft.XMLDOM");
       		xmlDom.async=false;
       		xmlDom.loadXML(xmlstring);
	}
	catch(err)
 	{
                throw("Error: Not able to create XMLDOM from string.  Make sure XML is well formed. ");
	}

        if(xmlDom.text=="" && xmlDom.xml == "")
                throw("Error: Not able to create XMLDOM from string.  Make sure XML is well formed. ");

   return xmlDom;
}

/** @ignore */
MLA.unescapeXMLCharEntities = function(stringtoconvert)
{
	var unescaped = "";
	unescaped = stringtoconvert.replace(/&amp;/g,"&");
	unescaped = unescaped.replace(/&lt;/g,  "<");
	unescaped = unescaped.replace(/&gt;/g,  ">");
	unescaped = unescaped.replace(/&quot;/g,"\"");
	unescaped = unescaped.replace(/&apos;/g,"\'");
	return unescaped;
}
/** @ignore */
MLA.escapeXMLCharEntities = function(stringtoconvert)
{
	var escaped = "";
	escaped = stringtoconvert.replace(/&/g,"&amp;");
	escaped = escaped.replace(/</g, "&lt;");
	escaped = escaped.replace(/>/g, "&gt;");
	escaped = escaped.replace(/\"/g,"&quot;");
	escaped = escaped.replace(/\'/g, "&apos;");
	return escaped;
}
/**
 * Returns ids for custom parts (not built-in) that are part of the active Open XML package. (.docx, .xlsx, .pptx, etc.)
 * @returns the ids for custom XML parts in active Open XML package as array of string
 * @type Array 
 */
MLA.getCustomXMLPartIds = function()
{ 
	var partIds = window.external.getCustomXMLPartIds();

	var errMsg = MLA.errorCheck(partIds);

	if(errMsg!=null)
	  throw("Error: Not able to get CustomXMLPartIds; "+errMsg);

	var customPartIds = partIds.split(" ");

	if(customPartIds.length ==1)
	{
		if (customPartIds[0] == null || customPartIds[0] == "")
		{
			customPartIds.length = 0;
		}
	}

	return customPartIds;
}

/**
 * Returns the custom XML part, identified by customXMLPartId, that is part of the active Open XML package. (.docx, .xlsx, .pptx, etc.)
 * @param customXMLPartId the id of the custom part to be fetched from the active package
 * @return the XML for the custom part as a DOM object. 
 * @type Microsoft.XMLDOM object 
 * @throws Exception if there is error retrieving the custom part 
 */
MLA.getCustomXMLPart = function(customXMLPartId)
{
	var customXMLPart = window.external.getCustomXMLPart(customXMLPartId);

	var errMsg = MLA.errorCheck(customXMLPart);

	if(errMsg!=null){
	   throw("Error: Not able to getCustomXMLPart; "+errMsg);
	}
        
	var v_cp;

	if(customXMLPart=="")
	{
		v_cp=null;
	}
        else
	{
        	v_cp = MLA.createXMLDOM(customXMLPart); 
	}

	return v_cp;
}

/** Adds custom part to active Open XML package.  Returns the id of the part added.
 *@param customPartXML Either A) an XMLDOM object that is the custom part to be added to the active Open XML package, or B)The string serialization of the XML to be added as a custom part to the active Open XML package. ( The XML must be well-formed. )
 *@return id for custom part added 
 *@type String
 *@throws Exception if unable to add custom part
 */
MLA.addCustomXMLPart = function(customPartXml)
{
	var v_customPart ="";
	if(customPartXml.xml)
	{
               v_customPart=customPartXml.xml;
	}
	else
	{
	       v_customPart=customPartXml
	}
	
	var newId = window.external.addCustomXMLPart(v_customPart);

	var errMsg = MLA.errorCheck(newId);
	if(errMsg!=null)
	   throw("Error: Not able to addCustomXMLPart; "+errMsg);

	if(newId =="")
	  newId=null;

	return newId;
}

/** Deletes custom part from Active Open XML package identified by id.
 *@param customXMLPartId the id of the custom part to be deleted from the active Open XML package.
 *@throws Exception if unable to delete custom part
 */
MLA.deleteCustomXMLPart = function(customXMLPartId)
{
	var deletedPart = window.external.deleteCustomXMLPart(customXMLPartId);

        var errMsg = MLA.errorCheck(deletedPart);
	if(errMsg!=null)
	   throw("Error: Not able to deleteCustomXMLPart; "+errMsg);
     
	if(deletedPart=="")
	  deletedPart = null;

}

/** @ignore */
MLA.isArray = function(obj)
{
 return obj.constructor == Array;
}

/**
 *
 *Returns MLA.config. The fields for this object are version, url, and theme.  
version - the version of the Addin library, url - the url used by the Addin WebBrowser control, theme - the current color scheme used by Office. 
 *@throws Exception if unable to create MLA.config object
 */
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
/** Inserts image into the Presentation.  
 *@param url a url to XQuery module that will return the image when evaluated  
 *@param user username for the MarkLogic Server the picuri connects with
 *@param pwd password for the MarkLogic Server the picuri connects with
 *@throws Exception if unable to insert image
 */
MLA.insertImage = function(url,user,pwd)
{
	
	var inserted = window.external.InsertImage(url,user,pwd);
	var errMsg = MLA.errorCheck(inserted);

	if(errMsg!=null)
	   throw("Error: Not able to insertImage; "+errMsg);

	return inserted;
}

/** Inserts slide, identified by slideIdx,  into the active presentation at current slide position.  
 *@param tmpPath the directory (including path) where the local copy of presentation will be saved.  
 *@param filename the name of the .pptx file
 *@param slideIdx the index of the slide within the source powerpoint file to be copied
 *@param url the url of the .pptx to be downloaded
 *@param user the username of the MarkLogic Server the url connects with
 *@param pwd the password of the MarkLogic Server the url connects with
 *@param retain true or false setting determines whether background style of copied slide will be retained when copied to active presentation
 *@throws Exception if unable to copy slide to active presentation 
 */
MLA.insertSlide = function(tmpPath, filename, slideidx, url, user, pwd,retain)
{
	var msg = window.external.insertSlide(tmpPath,filename,slideidx,url,user,pwd,retain);
	var errMsg = MLA.errorCheck(msg);
	if(errMsg!=null)
	   throw("Error: Not able to insertSlide; "+errMsg);

	return msg;
}

/**
 * Returns the path being used for the /temp dir on the client system.
 * @return /temp path on client system
 * @type string
 * @throws Exception if unable to retrieve the /temp path
 */
MLA.getTempPath = function()
{
	//alert("IN HERE");
	var msg=window.external.getTempPath();
	var errMsg = MLA.errorCheck(msg);

        if(errMsg!=null) 
        	throw("Error: Not able to getTempPath; "+errMsg);

	return msg;
}

/** Opens .pptx  into PowerPoint from local copy saved to client from MarkLogic.
 *@param tmpPath the directory (including path) where the local copy of presentation will be saved.  
 *@param docuri the uri of the .pptx within MarkLogic
 *@param url the url for fetching the .pptx to be downloaded
 *@param user the username for the MarkLogic Server the url connects with
 *@param pwd the password for the MarkLogic Server the url connects with
 *@throws Exception if unable to download and open local copy 
 */
MLA.openPPTX = function(tmpPath, docuri, url, user, pwd)
{
	//alert("tmpPath: "+tmpPath+" docuri: "+docuri+" url:"+url+" user/pwd:"+ user+pwd);
	var msg = window.external.openPPTX(tmpPath, docuri, url, user, pwd);
        var errMsg = MLA.errorCheck(msg);

        if(errMsg!=null) 
        	throw("Error: Not able to openPPTX; "+errMsg);

	return msg;
}
/** Inserts text into the Presentation at cursor position.  
 *@param text text to inser
 *@throws Exception if unable to insert text
 */
MLA.insertText = function(text)
{
	var msg = window.external.insertText(text);
        var errMsg = MLA.errorCheck(msg);

        if(errMsg!=null) 
        	throw("Error: Not able to insertText; "+errMsg);

	return msg;
}

/** Embeds OLE into the Presentation. the OLE is downloaded to client and saved into Presentation from local file.
 *@param tmpPath the directory (including path) where the local copy of object to be embedded will be saved.  
 *@param filename the name of the file to be embedded. tmpPath + filename should be the name and path of file on client.  
 *@param url the url for the file to be downloaded and embedded
 *@param user username for MarkLogic Server url connects with
 *@param pwd password for MarkLogic Server url connects with
 *@throws Exception if unable to embedOLE 
 */
MLA.embedOLE = function(tmpPath, title, url, usr, pwd)
{
	var msg = window.external.embedOLE(tmpPath, title, url, usr, pwd);
	var errMsg = MLA.errorCheck(msg);
        //alert("errMsg"+errMsg);
        if(errMsg!=null) 
        	throw("Error: Not able to embedOLE; "+errMsg);

	return msg;
}

/** Launches Windows form on client as simple SaveAs text box. 
 * @return text entered by user as filename into form
 * @type string 
 *@throws Exception if unable to return text 
 */
MLA.useSaveFileDialog =function()
{
	var msg=window.external.useSaveFileDialog();
	var errMsg = MLA.errorCheck(msg);

        if(errMsg!=null) 
        	throw("Error: Not able to embedOLE; "+errMsg);

	return msg;
}

/** Converts .pptx filename to image directory name.
 * @param filename the name to be converted
 * @return converted_name replaces .pptx of filename with _PNG 
 * @type string
 * @throws Exception if unable to retrieve the /temp path
 */
MLA.convertFilenameToImageDir = function(filename)
{
	//alert("IN HERE");
	var msg=window.external.convertFilenameToImageDir(filename);
	var errMsg = MLA.errorCheck(msg);

        if(errMsg!=null) 
        	throw("Error: Not able to convert filename to image directory name; "+errMsg);

	return msg;
}

/**
 * Returns the path being used for the active Presentation on the client system.
 * @return path path is path for where current Presentation is saved on client
 * @type string
 * @throws Exception if unable to retrieve the presentation path
 */
MLA.getPresentationPath = function()
{
	var msg=window.external.getPresentationPath();
	var errMsg = MLA.errorCheck(msg);

        if(errMsg!=null) 
        	throw("Error: Not able to getPresentationPath; "+errMsg);

	return msg;
}

/**
 * Returns the name being used for the active Presentation on the client system.
 * @return presentation_name the name of the active Presentation
 * @type string
 * @throws Exception if unable to retrieve the presentation name
 */
MLA.getPresentationName = function()
{
	var msg=window.external.getPresentationName();
	var errMsg = MLA.errorCheck(msg);

        if(errMsg!=null) 
        	throw("Error: Not able to getPresentationName; "+errMsg);

	return msg;
}

/**
 * Saves .pptx for active Presentation on the client system.
 * @param filename the filename (including path) to save Presentation as on client system
 * @throws Exception if unable to save local copy
 */
MLA.saveLocalCopy = function(filename)
{
	var msg = window.external.saveLocalCopy(filename);
        var errMsg = MLA.errorCheck(msg);

        if(errMsg!=null) 
        	throw("Error: Not able to saveLocalCopy; "+errMsg);

	return msg;
}


/** Saves active Presentation to MarkLogic from client system.  .pptx being saved to ML must already exist (be saved) on client and have both path and name.
 *@param filename the name of the file (including path) to be saved to MarkLogic  
 *@param url the url on MarkLogic that the client calls to upload the presentation
 *@param user username for MarkLogic Server url connects with
 *@param pwd password for MarkLogic Server url connects with
 *@throws Exception if unable to save the active Presentation
 */
MLA.saveActivePresentation = function(filename, url, user, pwd)
{

	var msg=window.external.saveActivePresentation(filename, url, user, pwd);
	var errMsg = MLA.errorCheck(msg);

        if(errMsg!=null) 
        	throw("Error: Not able to saveActivePresentation; "+errMsg);

	return msg;
}


/** Saves directory of images on client. saves this directory and its contents to MarkLogic.
 *@param imgdir the name of the directory (including path) where to save the images on the client system  
 *@param url the url on MarkLogic that the client calls to upload the image directory and contents to MarkLogic
 *@param user username for MarkLogic Server url connects with
 *@param pwd password for MarkLogic Server url connects with
 *@throws Exception if unable to save images
 */
MLA.saveImages = function(imgdir, url, user, pwd)
{
	var msg = window.external.saveImages(imgdir,url,user,pwd);
	var errMsg = MLA.errorCheck(msg);

        if(errMsg!=null) 
        	throw("Error: Not able to saveImages; "+errMsg);

	return msg;
}

/** Saves Presentation on client as .pptx. Saves same .pptx to MarkLogic. Saves images for Presentation to client.  Saves same images to MarkLogic.
 *@param saveasdir the name of the directory (including path) where to save the local copies
 *@param saveasname the name to save the .pptx as (no path)
 *@param url the url on MarkLogic that the client calls to upload the Presentation and images to MarkLogic
 *@param user username for MarkLogic Server url connects with
 *@param pwd password for MarkLogic Server url connects with
 *@throws Exception if unable to save active Presentation and images
 */
MLA.saveActivePresentationAndImages = function(saveasdir, saveasname, url, user, pwd)
{
	var msg=window.external.saveActivePresentationAndImages(saveasdir, saveasname, url, user, pwd);
	var errMsg = MLA.errorCheck(msg);

        if(errMsg!=null) 
        	throw("Error: Not able to saveActivePresentationAndImages; "+errMsg);

	return msg;
}

