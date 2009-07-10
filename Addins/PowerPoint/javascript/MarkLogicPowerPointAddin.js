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
 * Returns version of MarkLogicWordAddin.js library
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
/** Utility function for creating Microsoft.XMLDOM object from string
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


/** Returns the document.xml for the ActiveDocument as XMLDOM object. 
 * @return document.xml from the active Open XML package -- NOTE: This is the materialized view of the document.  If you have Content Controls bound to XML data islands within the .docx package, you'll only see the inline text, not the references, nor the original XML from which the value was mapped.
 * @type Microsoft.XMLDOM object
 * @throws Exception if unable to retrieve Styles.xml from the ActiveDocument.
 */
MLA.getActiveDocXml = function()
{
	var documentXml = window.external.getActiveDocXml();

        var errMsg = MLA.errorCheck(documentXml);
	if(errMsg!=null)
	   throw("Error: Not able to getActiveDocumentXml; "+errMsg);

	if(documentXml=="")
          documentXml=null;

	var v_documentXml = MLA.createXMLDOM(documentXml);
	return v_documentXml;

}


/** Inserts document.xml into the ActiveDocument package in Word.  This will replace the contents of the entire document the user is currently viewing. 
 *
 * As this only allows insert of document.xml, it is assumed that whatever references required by document.xml by other xml files in the package currently being authored (styles, themes, etc.) already exist.
 * 
 *@param documentXml this parameter may either be A) an XMLDOM object that is the XML equivalent of the document.xml to be inserted,or B) a String, that is the serialized, well-formed XML of the document.xml to be inserted.
 *@throws Exception if unable to set the documentXml
 */
MLA.setActiveDocXml = function(documentXml)
{
	var v_document="";

	if(documentXml.xml)
	{ 
               v_document=documentXml.xml;
	}
	else
	{ 
	       v_document = documentXml;
	}

        var inserted = window.external.setActiveDocXml(v_document);

	var errMsg = MLA.errorCheck(inserted);
	if(errMsg!=null)
	   throw("Error: Not able to setActiveDocXml. Make sure XML is well-formed and valid wordprocessingML; "+errMsg);

	if(inserted=="")
	  inserted = null;
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
/** Inserts image into the ActiveDocument at current cursor position.  
 *@param picurl a url to an XQuery module that will return the image when evaluated.  
 *@param uname username for Server
 *@param pwd password for Server
 *@throws Exception if unable to insert text
 */
MLA.insertImage = function(picuri,uname,pwd)
{
	
	var inserted = window.external.InsertImage(picuri,uname,pwd);
	var errMsg = MLA.errorCheck(inserted);
	if(errMsg!=null)
	   throw("Error: Not able to insertImage; "+errMsg);

}

/** Inserts slide, identified by slideIdx,  into the active presentation at current slide position.  
 *@param tmpPath the directory where the local copy will be saved.  
 *@param filename the name of the powerpoint file
 *@param slideIdx the index of the slide within the source powerpoint file to be copied
 *@param url the url of the .pptx to be downloaded
 *@param user the username of the MarkLogic Server the url connects with
 *@param pwd the password of the MarkLogic Server the url connects with
 *@param retain true or false setting determines whether background style of copied slide will be retained when copied to active presentation
 *@throws Exception if unable to copy slide to active presentation 
 */
MLA.copyPasteSlideToActive = function(tmpPath, filename,slideidx, url, user, pwd,retain)
{
	//alert("IN MLA2 tmpPath: "+tmpPath+" fileanme: "+filename +"slidenumber"+slideidx+" url: "+url +" user/pwd"+user+"|"+pwd);
	////master may differ, takes first master, need to insure its correct one in function in AddIn

	var msg = window.external.copyPasteSlideToActive(tmpPath,filename,slideidx,url,user,pwd,retain);
	var errMsg = MLA.errorCheck(msg);
	if(errMsg!=null)
	   throw("Error: Not able to copyPasteSlideToActive; "+errMsg);

	return msg;
}

/**
 * Returns the path being used for the /temp dir on the client system
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

/** opens .pptx by downloading local copy to client  
 *@param tmpPath the directory where the local copy will be saved.  
 *@param docuri the uri of the .pptx within MarkLogic
 *@param url the url of the .pptx to be downloaded
 *@param user the username of the MarkLogic Server the url connects with
 *@param pwd the password of the MarkLogic Server the url connects with
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
