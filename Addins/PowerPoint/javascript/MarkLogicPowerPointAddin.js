/* 
Copyright 2009-2010 Mark Logic Corporation

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
//import json2.js
/** @ignore */
if (typeof(ml_scripts) == 'undefined') var ml_scripts = new Object();
/** @ignore */
ml_js_import('js/json2.js');

/** @ignore */
function ml_js_import(jsFile) {
         if (ml_scripts[jsFile] != null) return;
         var scriptElt = document.createElement('script');
             scriptElt.type = 'text/javascript';
             scriptElt.src = jsFile;
         document.getElementsByTagName('head')[0].appendChild(scriptElt);
         ml_scripts[jsFile] = jsFile; 
}
/** 
 * @fileoverview  API documentation for MarkLogicPowerPointAddin.js
 *
 *
 *
 * {@link http://www.marklogic.com} 
 *
 * @author Pete Aven pete.aven@marklogic.com
 * @version 1.1-1
 */

/**
 * The MLA namespace is used for global attribution. The methods within this namespace provide ways of interacting with an active Open XML document through a WebBrowser control. The control must be deployed within an Addin in Office 2007.
 *
 * The functions here provide ways for interacting with the active presentation in PowerPoint ; however, the functions getCustomXMLPart(), getCustomXMLPartIds(), addCustomXMLPart(), and deleteCustomXMLPart() will work for any Open XML package, provided they are used within the context of an Addin for the appropriate Office application.
 */
var MLA = {};
/*  following left here for using jsdocs
 *  comment above and use so jsdocs generate */
/*
function MLA(){
      this.getClassName = function(){
      return "MLA";
   
      }
}
*/
/** @ignore */
MLA.version = { "release" : "1.1-1" };

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
 * @type string
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
 *@type string
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
 *@type string
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
 *@type string
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
 *@type string
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
 *@type string
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
 *@type string
 *@throws Exception if unable to embedOLE 
 */
MLA.embedOLE = function(tmpPath, title, url, usr, pwd)
{
	var msg = window.external.embedOLE(tmpPath, title, url, usr, pwd);
	var errMsg = MLA.errorCheck(msg);

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
 * @type string
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
 *@type string
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
 *@type string
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
 *@type string
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

/** Inserts JSON string as table in Active Presentation.  
 * see template example in insertJSONTable() found in Samples/officesearch/officesearch.js for required JSON format.
 *@param table the JSON representation of the table to be inserted
 *@type string
 *@throws Exception if unable to save insert table into Presentation
 */
MLA.insertJSONTable = function(table)
{
	var msg=window.external.insertJSONTable(table);
	var errMsg = MLA.errorCheck(msg);

        if(errMsg!=null) 
        	throw("Error: Not able to insertJSONTable(); "+errMsg);

	return msg;
}

/* ================== BEGIN Additions for PPT TK UPDATE 1.1-1  ==================================== */

/** Returns slide name of active slide.
 * @return slideName
 * @type string
 * @throws Exception if unable to retrieve the slide name
 */
MLA.getSlideName = function()
{
	var msg = window.external.getSlideName();
	var errMsg = MLA.errorCheck(msg);

        if(errMsg!=null) 
        	throw("Error: Not able to getSlideName(); "+errMsg);

	return msg;
}

/** Returns index of active slide.
 * @return slideIndex 
 * @type string
 * @throws Exception if unable to retrieve the slide index
 */
MLA.getSlideIndex = function()
{
	var msg = window.external.getSlideIndex();
	var errMsg = MLA.errorCheck(msg);

        if(errMsg!=null) 
        	throw("Error: Not able to getSlideIndex(); "+errMsg);

	return msg;
}

/** Returns count of slides for active presentation.
 * @return slideCount 
 * @type string
 * @throws Exception if unable to retrieve the slide count
 */
MLA.getPresentationSlideCount = function()
{
	var msg = window.external.getPresentationSlideCount();
	var errMsg = MLA.errorCheck(msg);

        if(errMsg!=null) 
        	throw("Error: Not able to getPresentationSlideCount(); "+errMsg);

	return msg;
}

/** Adds tag to active presentation
 * @param tagName the name of the tag
 * @param tagValue the value of the tag
 * @type string
 * @throws Exception if unable to add tag to presentation
*/
MLA.addPresentationTag = function(tagName, tagValue)
{ 
	var msg = window.external.addPresentationTag(tagName,tagValue);

	var errMsg = MLA.errorCheck(msg);

        if(errMsg!=null) 
        	throw("Error: Not able to addPresentationTag(); "+errMsg);

	return msg;
}

/** Deletes tag from active presentation.
 * @param tagName the name of the tag to be deleted
 * @type string
 * @throws Exception if unable to delete the tag from the presentation
 */
MLA.deletePresentationTag = function(tagName)
{
        var msg = window.external.deletePresentationTag(tagName);
	var errMsg = MLA.errorCheck(msg);

        if(errMsg!=null) 
        	throw("Error: Not able to deletePresentationTag(); "+errMsg);

	return msg;

}

/** Adds tag to slide.
 * @param slideIndex the index of the slide to be tagged
 * @param tagName the name of the tag
 * @param tagValue the value of the tag
 * @type string
 * @throws Exception if unable to add tag to slide 
 */
MLA.addSlideTag = function(slideIndex, tagName, tagValue)
{
	var msg = window.external.addSlideTag(slideIndex, tagName, tagValue);
	var errMsg = MLA.errorCheck(msg);

        if(errMsg!=null) 
        	throw("Error: Not able to addSlideTag(); "+errMsg);

	return msg;
}

/** Deletes tag from slide.
 * @param slideIndex the index of the slide
 * @param tagName the name of the tag to be deleted
 * @type string
 * @throws Exception if unable to delete slide tag 
 */
MLA.deleteSlideTag = function(slideIndex, tagName)
{
	var msg = window.external.deleteSlideTag(slideIndex, tagName);
	var errMsg = MLA.errorCheck(msg);

        if(errMsg!=null) 
        	throw("Error: Not able to deleteSlideTag(); "+errMsg);

	return msg;
}

/** Adds tag to shape.
 * @param slideIndex the index of the slide containing the shape 
 * @param shapeName the name of the shape to be tagged 
 * @param tagName the name of the tag
 * @param tagValue the value of the tag
 * @type string
 * @throws Exception if unable to add shape tag
 */
MLA.addShapeTag = function(slideIndex, shapeName, tagName, tagValue)
{
	var msg = window.external.addShapeTag(slideIndex, shapeName, tagName, tagValue);
	var errMsg = MLA.errorCheck(msg);

        if(errMsg!=null) 
        	throw("Error: Not able to addShapeTag(); "+errMsg);

	return msg;

}

/** Deletes tag from shape.
 * @param slideIndex the index of the slide containing the shape 
 * @param shapeName the name of the tagged shape 
 * @param tagName the name of the tag to be deleted
 * @type string
 * @throws Exception if unable to delete shape tag
 */
MLA.deleteShapeTag = function(slideIndex, shapeName, tagName)
{
	var msg = window.external.deleteShapeTag(slideIndex, shapeName, tagName);
	var errMsg = MLA.errorCheck(msg);

        if(errMsg!=null) 
        	throw("Error: Not able to deleteShapeTag(); "+errMsg);

	return msg;
}

/** Adds tags (tag names and associated values) to shape from jsonSerialization of tags.  You can get the json serialization of tags by using MLA.jsonStringify(shapeRangeView.tags)
 * @param slideIndex the index of the slide containing the shape 
 * @param shapeName the name of the shape to be tagged 
 * @param jsonTags the tags (name, value) to be added to the shape
 * @type string
 * @throws Exception if unable to add shape tags 
 */
MLA.addShapeTags = function(slideIndex, shapeName, jsonTags)
{
 	var msg = window.external.addShapeTags(slideIndex, shapeName, jsonTags);
	var errMsg = MLA.errorCheck(msg);

        if(errMsg!=null) 
        	throw("Error: Not able to addShapeTags(); "+errMsg);

	return msg;
}

/** Adds tags (tag names and associated values) to slide from jsonSerialization of tags.  You can get the json serialization of tags by using MLA.jsonStringify(shapeRangeView.tags) or MLA.jsonStringify(MLA.getSlideTags()). Most likely, you'd save the serialization in a custom part and apply to new presentations when reusing slide.
 * @param slideIndex the index of the slide containing the shape 
 * @param jsonTags the tags (name, value) to be added to the slide
 * @type string
 * @throws Exception if unable to add slide tags 
 */
MLA.addSlideTags = function(slideIndex, jsonTags)
{
 	var msg = window.external.addSlideTags(slideIndex, jsonTags);
	var errMsg = MLA.errorCheck(msg);

        if(errMsg!=null) 
        	throw("Error: Not able to addSlideTags(); "+errMsg);

	return msg;
}

/** Adds tags (tag names and associated values) to active presentation from jsonSerialization of tags.  You can get the json serialization of tags by using MLA.jsonStringify(MLA.getPresentationTags()).  Most likely, you'd save the serialization in a custom XML part and apply to new presentations based on some business logic.
 * @param jsonTags the tags (name, value) to be added to the presentation 
 * @type string
 * @throws Exception if unable to add presentation tags 
 */
MLA.addPresentationTags = function(jsonTags)
{
 	var msg = window.external.addPresentationTags(jsonTags);
	var errMsg = MLA.errorCheck(msg);

        if(errMsg!=null) 
        	throw("Error: Not able to addPresentationTags(); "+errMsg);

	return msg;
}

//adds tag to selected range where addShapeTag adds tag to single shape, that may or may not be currently selected
/** Adds tag to all selected shapes in the active slide
 * @param tagName the name of the tag
 * @param tagValue the value of the tag 
 * @type string
 * @throws Exception if unable to add shape range tags
 */
MLA.addShapeRangeTag = function(tagName, tagValue)
{
	var msg = window.external.addShapeRangeTag(tagName, tagValue);
	var errMsg = MLA.errorCheck(msg);

        if(errMsg!=null) 
        	throw("Error: Not able to addShapeRangeTag(); "+errMsg);

	return msg;	
}

/**
 * Returns the count of how many shapes are currently selected in the active slide.
 * @return shapeCount the number of currently selected shapes
 * @type string
 * @throws Exception if unable to get shape range count
 */
MLA.getShapeRangeCount = function()
{
	var msg = window.external.getShapeRangeCount();
	var errMsg = MLA.errorCheck(msg);

        if(errMsg!=null) 
        	throw("Error: Not able to getShapeRangeCount(); "+errMsg);

	return msg;	
}

/**
 * Returns the name of the currently selected shape.
 * @return shapeName the name of the shape
 * @type string
 * @throws Exception if unable to retrieve the shape range name
 */
MLA.getShapeRangeName = function()
{
	var msg = window.external.getShapeRangeName();
	var errMsg = MLA.errorCheck(msg);

        if(errMsg!=null) 
        	throw("Error: Not able to getShapeRangeName(); "+errMsg);

	return msg;
}
/**
 * Sets the name being used for the active Presentation on the client system.
 * @param slideIndex the index of the slide containing the shape
 * @param origName the original name of the shape
 * @param newName the name to set for the shape
 * @type string
 * @throws Exception if unable to set shape range name
 */
MLA.setShapeRangeName = function(slideIndex, origName,newName)
{
	var msg = window.external.setShapeRangeName(slideIndex, origName, newName);
	var errMsg = MLA.errorCheck(msg);

        if(errMsg!=null) 
        	throw("Error: Not able to setShapeRangeName(); "+errMsg);

	return msg;
}

//no different than getShapeRangeName, just returns all selected instead of one
/**
 * Gets the names of all currently selected shapes in the active slide.
 * @return then names of all currently selected shapes
 * @type Array 
 * @throws Exception if unable to get shape range name
 */
MLA.getShapeRangeShapeNames = function()
{
	var msg = window.external.getShapeRangeShapeNames();
	var errMsg = MLA.errorCheck(msg);

	var tokens = msg.split("|");

        if(errMsg!=null) 
        	throw("Error: Not able to getShapeRangeShapeNames(); "+errMsg);

	return tokens;
}

/**
 * Gets the names of all shapes for the slide specified by slideIndex.
 * @param slideIndex the index of the slide
 * @return then names of all shapes on the slide
 * @type Array 
 * @throws Exception if unable to get slide shape names
 */
MLA.getSlideShapeNames = function(slideIndex)
{
	var msg = window.external.getSlideShapeNames(slideIndex);
	var errMsg = MLA.errorCheck(msg);

	var tokens = msg.split("|");

        if(errMsg!=null) 
        	throw("Error: Not able to getShapeRangeShapeNames(); "+errMsg);

	return tokens;

}

/** @ignore */
MLA.getShapeRangeInfoORIG = function(shapeName)
{
	var msg = window.external.getShapeRangeInfo(shapeName);
	var errMsg = MLA.errorCheck(msg);

        if(errMsg!=null) 
        	throw("Error: Not able to getShapeRangeInfo(); "+errMsg);

	//var shape =  eval('('+msg+')');
	var shape =  MLA.jsonParse(msg);
	
	return shape;
}

/**
 * Gets  a shapeRangeView object
 * @param slideIndex the index of the slide containing the shape
 * @param shapeName the name of the shape
 * @return ShapeRangeVIew
 * @type ShapeRangeView
 * @throws Exception if unable to get shape range view
 */
MLA.getShapeRangeView = function(slideIndex, shapeName)
{
	var msg = window.external.getShapeRangeView(slideIndex, shapeName);
	var errMsg = MLA.errorCheck(msg);

	//alert("msg");

        if(errMsg!=null) 
        	throw("Error: Not able to getShapeRangeView(); "+errMsg);

	try{ 
		//alert("JSON: "+msg);
         	//var tmpshape =  eval('('+msg+')');
         	var tmpshape = MLA.jsonParse(msg);
		//alert(tmpshape.paragraphs.length);
	}catch(err)
	{ 
		alert("ERROR: " +err.description);
	}


	var shapeTags = tmpshape.tags;
	var shapeParas = tmpshape.paragraphs;
	var picFormat = tmpshape.pictureFormat;

	delete tmpshape.tags;
	delete tmpshape.paragraphs;
	delete tmpshape.pictureFormat;

	var shapeRangeView = new MLA.ShapeRangeView();
	shapeRangeView.shape = tmpshape;
	shapeRangeView.tags = shapeTags;
	shapeRangeView.paragraphs = shapeParas;
	shapeRangeView.pictureFormat = picFormat;
	
	return shapeRangeView;
}

/**
 * Gets the json serialization of tags for the active presentation.  
 * @return the json serialization of tags for the active presentation. 
 * @type string
 * @throws Exception if unable to get presentation tags
 */
MLA.getPresentationTags = function()
{
	var msg = window.external.getPresentationTags();
	var errMsg = MLA.errorCheck(msg);

        if(errMsg!=null) 
        	throw("Error: Not able to getPresentationTags(); "+errMsg);

	//var tags = eval('('+msg+')');
	var tags = MLA.jsonParse(msg);

	return tags;

}

/**
 * Gets the json serialization of tags for a slide in the active presentation.  
 * @param slideIndex the index of the slide
 * @return json serialization of tags for a slide
 * @type string
 * @throws Exception if unable to get slide tags
 */
MLA.getSlideTags = function(slideIndex)
{
	var msg = window.external.getSlideTags(slideIndex);
	var errMsg = MLA.errorCheck(msg);

        if(errMsg!=null) 
        	throw("Error: Not able to getSlideTags(); "+errMsg);

	//var tags = eval('('+msg+')');
	var tags = MLA.jsonParse(msg);

	return tags;
}

/**
 * Add a shape to the active presentation
 * @param slideIndex the index of the slide to add the shape to
 * @param shapeRangeView the json serialization of a ShapeRangeView object.
 * @return shapeName the name of the new shape 
 * @type string
 * @throws Exception if unable to add shape
 */
MLA.addShape = function(slideIndex, shapeRangeView)
{
	var msg="";
	try
	{
	  var shape = MLA.jsonStringify(shapeRangeView.shape);
	  var tags = MLA.jsonStringify(shapeRangeView.tags);
	  var paragraphs = MLA.jsonGetParagraphs(shapeRangeView.paragraphs);

          msg=window.external.addShape(slideIndex,shape, tags, paragraphs);

	  var errMsg = MLA.errorCheck(msg);

          if(errMsg!=null) 
        	throw("Error: Not able to addShape(); "+errMsg);

	}catch(err)
	{
		throw("Error: Not able to addShape() JS; "+ err.description);
	}

	return msg;


}

/**
 * Adds a slide to the active presentation.  
 * @param slideIndex the index of where to add the slide
 * @param customLayout the layout to use for the added slide
 * @type string
 * @throws Exception if unable to add slide 
 */
MLA.addSlide = function(slideIndex, customLayout)
{
  	var msg = window.external.addSlide(slideIndex, customLayout);
	var errMsg = MLA.errorCheck(msg);

        if(errMsg!=null) 
        	throw("Error: Not able to addSlide(); "+errMsg);

	return msg; //should return slide Index or Name
}

/**
 * Deletes a slide from the active presentation.  
 * @param slideIndex the index of slide to delete
 * @type string
 * @throws Exception if unable to delete slide 
 */
MLA.deleteSlide = function(slideIndex)
{
  	var msg = window.external.deleteSlide(slideIndex);
	var errMsg = MLA.errorCheck(msg);

        if(errMsg!=null) 
        	throw("Error: Not able to deleteSlide(); "+errMsg);

	return msg;
}

/**
 * Deletes a shape from the active slide in the active presentation.  
 * @param slideIndex the index of slide containing shape to be deleted
 * @param shapeName the name of the shape to be deleted
 * @type string
 * @throws Exception if unable to delete shape 
 */
MLA.deleteShape = function(slideIndex, shapeName)
{
	var msg = window.external.deleteShape(slideIndex, shapeName);
	var errMsg = MLA.errorCheck(msg);

        if(errMsg!=null) 
        	throw("Error: Not able to deleteShape(); "+errMsg);

	return msg;
}

/** @ignore */
MLA.jsonGetParagraphs =  function(paragraphs)
{
        var jsonPara = "{";
	jsonPara += "\"paragraphAlignment\": [";

        for(var j =0;j<paragraphs.length;j++)
	{
		var para = paragraphs[j];
		jsonPara += "\""+para.paragraphAlignment+"\",";
        }

	if(paragraphs.length>=1)
	  jsonPara = jsonPara.substring(0,jsonPara.length-1);

	jsonPara += "],";

	//
	jsonPara +=  "\"paragraphBulletType\": [";

        for(var j =0;j<paragraphs.length;j++)
	{
		var para = paragraphs[j];
		jsonPara += "\""+para.paragraphBulletType+"\",";
        }

	if(paragraphs.length>=1)
	  jsonPara = jsonPara.substring(0,jsonPara.length-1);

	jsonPara += "],";
	//

	jsonPara += "\"runs\": [";

	var runLength;
	for(var k=0; k<paragraphs.length;k++)
	{
		var para = paragraphs[k];
                
		runLength=para.runs.length;
		for(var l =0;l<para.runs.length;l++)
	        {
			var run = para.runs[l];
		        jsonPara+="["+"\""+k+"\",\""+run.fontName+"\",\""+
				                     run.fontSize+"\",\""+
						     run.fontRGB+"\",\""+
						     run.fontItalic +"\",\""+
						     run.fontUnderline +"\",\""+
						     run.fontBold +"\",\""+
						     run.text+"\"],";
		}
	
	}
	if(runLength>=1)
            jsonPara = jsonPara.substring(0,jsonPara.length-1);

	jsonPara +="]}";

	return jsonPara;

}

/**
 * Serializes an object as JSON.  
 * @param jsObject the object to be serialized as JSON
 * @param replacer optional replacer parameter
 * @type string
 * @throws Exception if unable to stringify
 */
MLA.jsonStringify = function(jsObject, replacer)
{
	var jsonString;

	try
	{
            jsonString = JSON.stringify(jsObject)
	}
	catch(err)
	{
		throw("Error: Not able to stringify(); "+err.description);
	}

	return jsonString;

}

/**
 * Deserializes a JSON string into an object.
 * @param jsString the JSON string to convert to object.
 * @param reviver optional reviver parameter
 * @type string
 * @throws Exception if unable to parse
 */
MLA.jsonParse = function(jsonString, reviver)
{
	var jsonObj;

	try
	{
            jsonObj = JSON.parse(jsonString)
	}
	catch(err)
	{
		throw("Error: Not able to parse(); "+err.description);
	}

	return jsonObj;
}

/**
 * Sets the picture attributes for an inserted msoPicture shape.
 * @param slideIndex the index of the slide containing the picture
 * @param shapeName the name of the picture shape
 * @param jsonPicFormat the JSON serialization of picture format, available by using MLA.jsonStringify(shapeRangeView.pictureFormat)
 * @type string
 * @throws Exception if unable to set picture format 
 */
MLA.setPictureFormat = function(slideIndex, shapeName, jsonPicFormat)
{
         
     msg=window.external.setPictureFormat(slideIndex, shapeName, jsonPicFormat);
     var errMsg = MLA.errorCheck(msg);

     if(errMsg!=null) 
         throw("Error: Not able to addShape(); "+errMsg);

     return msg;
}

/**
 * Create a new ShapeRangeView instance. 
 * @class A basic ShapeRangeView class.
 */
MLA.ShapeRangeView = function()
{
	var shape;
	var paragraphs;
	var tags;
	var pictureFormat;
}

