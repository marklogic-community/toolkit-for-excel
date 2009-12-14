/* 
Copyright 2008-2009 Mark Logic Corporation

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.

MarkLogicWordAddin.js - javascript api for interacting with webBrowser control within Custom Task Pane enabled in Word.
*/

/** 
 * @fileoverview  API documentation for MarkLogicWordAddin.js
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
 * The functions here provide ways for interacting with the ActiveDocument in Word; however, the functions getCustomXMLPart(), getCustomXMLPartIds(), addCustomXMLPart(), and deleteCustomXMLPart() will work for any Open XML package, provided they are used within the context of an Addin for the appropriate Office application.
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
 * @param xmlString the string to be loaded into a XMLDOM object.  The string must be serialized, well-formed XML.
 * @return Microsoft.XMLDOM object
 * @throws Exception if unable to create the XMLDOM object
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
/** Utility function to create a default WordprocessingML paragraph <w:p>, with no styles, for a given string.
 *
 * @param textstring the string to be converted into a WordprocessingML paragraph
 * @return Microsoft.XMLDOM object that is a WordprocessingML paragraph
 * @throws Exception if unable to create the paragraph
 */
MLA.createParagraph = function(textstring)
{ 
	var cleanstring = MLA.escapeXMLCharEntities(textstring);
	var newParagraphString = "<w:p xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'><w:r><w:t>"+
		                    cleanstring+
			         "</w:t></w:r></w:p>";
	var newPara = MLA.createXMLDOM(newParagraphString);
	return newPara;
}
/** Inserts text into the ActiveDocument at current cursor position.  Text will be styled according to whatever is the currently chosen style for text within Word. 
 * @param textToInsert the text to be inserted at the current cursor position in the ActiveDocument.
 * @throws Exception if unable to insert text 
 */
MLA.insertText = function(textToInsert)
{
	var textAdded = window.external.insertText(textToInsert);
	var errMsg = MLA.errorCheck(textAdded);
	if(errMsg!=null)
	   throw("Error: Not able to insert text "+errMsg);
     
	if(textAdded=="")
	  textAdded = null;
}
/**
 *  Returns the text currently selected (highlighted) by the user in the ActiveDocument.  Whatever is highlighted by the user will be returned in this function as a string.  A user may highlight text from multiple paragraphs or from tables.  For this reason, the function returns an array, where each element of the array is a string that contains the text contained by the structure highlighted by the user in the ActiveDocument.  The order of elements in the array represents the order of items that are highlighted in the ActiveDocument.
 * @return the text selected by the user in the ActiveDocument as strings. If nothing is selected, an empty array is returned.
 * @type Array
 * @param delimiter (optional) Text from tables will be returned as a single string, with cells delimited by tabs (default).  This param may be used to assign a different delimiter.  Note: there is no delimiter for text from paragraphs, as each paragraph will be captured in separate array elements.
 * @throws Exception if unable to retrieve the Selection
 */
MLA.getSelectionText = function(delimiter)
{
        if(delimiter == null) 
	    delimiter = "";

	var arrCount=0;
	var selCount =0;
	var selections = new Array();
        
	var selection =  window.external.getSelectionText(selCount,delimiter);

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
	  selection = window.external.getSelectionText(selCount,delimiter);


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



	if(selections.length == 1) 
	{
        	if (selections[0] == null ||selections[0].text == "")
		{
			selections.length = 0;
		}
	}

	return selections;
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
 * @param customPartXML Either A) an XMLDOM object that is the custom part to be added to the active Open XML package, or B)The string serialization of the XML to be added as a custom part to the active Open XML package. ( The XML must be well-formed. )
 * @return id for custom part added 
 * @type String
 * @throws Exception if unable to add custom part
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
 * @param customXMLPartId the id of the custom part to be deleted from the active Open XML package.
 * @return void
 * @type void
 * @throws Exception if unable to delete custom part
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

/** Returns the XML that represents what is currently selected (highlighted) by the user in the ActiveDocument as an XMLDOM object.  Whatever is highlighted by the user will be returned in this function as a block level element.  A user may highlight text that will be materialized as multiple sibling block elements in the XML.  For this reason, the function returns an array, where each element of the array is an XMLDOM object that contains the XML for the blocks highlighted by the user in the ActiveDocument.  The order of elements in the array represents the order of items that are highlighted in the ActiveDocument.
 * @return the blocks of XML currently selected by the user in the ActiveDocument as XMLDOM objects. If nothing is selected, an empty array is returned.
 * @type Array
 * @throws Exception if unable to retrieve the Selection
*/
MLA.getSelection = function()
{
	var arrCount=0;
	var selCount =0;
	var selections = new Array();
	var domSelections = new Array();
        
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


        if(!(selections[0]==null || selections[0]==""))
	{
		for(i=0;i<selections.length;i++)
		{
         		domSelections[i] = MLA.createXMLDOM(selections[i]);
		}
	
		if(domSelections.length == 1) 
		{
        		if (domSelections[0] == null ||domSelections[0].text == "")
			{
				domSelections.length = 0;
			}
		}
	}else
	{ 
		domSelections.length = 0; 
	}

	return domSelections;
}

/** Returns the final XML block that represents the Sentence at the current cursor position.  Nothing is required to be highlighted in the ActiveDocument.  If a selection is highlighted, this returns the XML for the sentence at the the current cursor position.  If there is no selection, the XML for the sentence at the current cursor position is still returned. 
 * @return the block of XML, as XMLDOM object,  for the Sentence immediately preceding the cursor position.
 * @type Microsoft.XMLDOM object
 * @throws Exception if unable to retrieve the XML for the Sentence.
*/
MLA.getSentenceAtCursor = function()
{
	var rangeXml = window.external.getSentenceAtCursor();

	var errMsg = MLA.errorCheck(rangeXml);
	if(errMsg!=null) 
	   throw("Error: Not able to get Sentence at Cursor; "+errMsg);

	var v_rangeXml = MLA.createXMLDOM(rangeXml);
	return v_rangeXml;
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

/** Returns the Styles.xml for the ActiveDocument as XMLDOM object. 
 * @return Styles.xml if no Styles.xml present in ActiveDocument package, returns null.
 * @type Microsoft.XMLDOM object
 * @throws Exception if unable to retrieve Styles.xml from the ActiveDocument.
 */
MLA.getActiveDocStylesXml = function()
{ 
	var stylesXml = window.external.getActiveDocStylesXml();

        var errMsg = MLA.errorCheck(stylesXml);
	if(errMsg!=null)
	   throw("Error: Not able to getActiveDocStylesXml; "+errMsg);

	if(stylesXml=="")
          stylesXml=null;

	var v_stylesXml = MLA.createXMLDOM(stylesXml);
	return v_stylesXml;
}

/** Inserts document.xml into the ActiveDocument package in Word.  This will replace the contents of the entire document the user is currently viewing. 
 *
 * As this only allows insert of document.xml, it is assumed that whatever references required by document.xml by other xml files in the package currently being authored (styles, themes, etc.) already exist.
 * 
 * @param documentXml this parameter may either be A) an XMLDOM object that is the XML equivalent of the document.xml to be inserted,or B) a String, that is the serialized, well-formed XML of the document.xml to be inserted.
 * @throws Exception if unable to set the documentXml
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

/** Inserts block-level xml into the ActiveDocument in Word at the current cursor position, or over the selected range (if selected).  If stylesxml parameter is defined, the blockxml will be inserted and the Styles.xml will be overwritten with the stylesxml contents. If stylesxml is not defined, then only the blockxml will be inserted ; the assumption being that whatever styles required by the blockxml are already present in the .docx Styles.xml.
 *
 * There are two main levels of content in the document.xml; block-level and inline content. Block level describes the structure of the document, and includes paragraphs and tables. Anything that can be inserted under <w:body> may be inserted here
 * 
 * @param blockXml this parameter may either be A) and XMLDOM object that is the block-level XML to be inserted,or an array of such objects,or B) a String, that is the serialized, well-formed XML of the block to be inserted, or an Array of such Strings.
 * @param stylesXml (optional) this parameter is either A) an XMLDOM object that contains Styles.xml for the pacakge, or B)a String, which is the serialized, well-formed XML that represents Styles.xml for the ActiveDocument package in Word.
 * @throws Exception if unable to insert the blockContent or Styles.xml
 */
MLA.insertBlockContent = function(blockContentXml,stylesXml)
{
	if(stylesXml == null) 
	    stylesXml = "";
  
       
	var v_block="";
	var v_styles="";

	if(blockContentXml.xml)
	{
               v_block=blockContentXml.xml;
	}
	else
	{
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
	}

	if(stylesXml.xml)
	{
             v_styles=stylesXml.xml;
	}
	else
	{
             v_styles=stylesXml;
	}
	
        var inserted = window.external.insertBlockContent(v_block,v_styles);

	var errMsg = MLA.errorCheck(inserted);
	if(errMsg!=null)
	   throw("Error: Not able to insertBlockContent; "+errMsg);

	if(inserted=="")
	  inserted = null;

}

/**
 *
 * Returns MLA.config. The fields for this object are version, url, and theme.  
version - the version of the Addin library, url - the url used by the Addin WebBrowser control, theme - the current color scheme used by Office. 
 * @throws Exception if unable to create MLA.config object
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
/**
 * Inserts image into the ActiveDocument at current cursor position.  
 * @param picurl a url to an XQuery module that will return the image when evaluated.  
 * @param uname username for Server
 * @param pwd password for Server
 * @return void
 * @type void
 * @throws Exception if unable to insert text
 */
MLA.insertImage = function(picuri,uname,pwd)
{
	
	var inserted = window.external.InsertImage(picuri,uname,pwd);
	var errMsg = MLA.errorCheck(inserted);
	if(errMsg!=null)
	   throw("Error: Not able to insertImage; "+errMsg);

	return inserted;

}

/**
 * Returns the path being used for the /temp dir on the client system.
 * @return /temp path on client system
 * @type String
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

/** Opens .docx  into Word from local copy saved to client from MarkLogic.
 * @param tmpPath the directory (including path) where the local copy of document will be saved.  
 * @param docuri the uri of the .docx within MarkLogic
 * @param url the url for fetching the .docx to be downloaded
 * @param user the username for the MarkLogic Server the url connects with
 * @param pwd the password for the MarkLogic Server the url connects with
 * @return void
 * @type void
 * @throws Exception if unable to download and open local copy 
 */
MLA.openDOCX = function(tmpPath, docuri, url, user, pwd)
{
	//alert("tmpPath: "+tmpPath+" docuri: "+docuri+" url:"+url+" user/pwd:"+ user+pwd);
	var msg = window.external.openDOCX(tmpPath, docuri, url, user, pwd);
        var errMsg = MLA.errorCheck(msg);

        if(errMsg!=null) 
        	throw("Error: Not able to openDOCX; "+errMsg);

	return msg;
}

/** Embeds OLE into the Document. the OLE is downloaded to client and saved into Document from local file.
 * @param tmpPath the directory (including path) where the local copy of object to be embedded will be saved.  
 * @param filename the name of the file to be embedded. tmpPath + filename should be the name and path of file on client.  
 * @param url the url for the file to be downloaded and embedded
 * @param user username for MarkLogic Server url connects with
 * @param pwd password for MarkLogic Server url connects with
 * @return void
 * @type void
 * @throws Exception if unable to embedOLE 
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

/**
 * Returns the path being used for the active document on the client system.
 * @return path path is path for where current document is saved on client
 * @type String
 * @throws Exception if unable to retrieve the document path
 */
MLA.getDocumentPath = function()
{
	var msg=window.external.getDocumentPath();
	var errMsg = MLA.errorCheck(msg);

        if(errMsg!=null) 
        	throw("Error: Not able to getDocumentPath; "+errMsg);

	return msg;
}

/**
 * Returns the name being used for the active Document on the client system.
 * @return document_name the name of the active Document
 * @type String
 * @throws Exception if unable to retrieve the Document name
 */
MLA.getDocumentName = function()
{
	var msg=window.external.getDocumentName();
	var errMsg = MLA.errorCheck(msg);

        if(errMsg!=null) 
        	throw("Error: Not able to getDocumentName; "+errMsg);

	return msg;
}

/**
 * Saves .docx for active Document on the client system.
 * @param filename the filename (including path) to save Document as on client system
 * @return void
 * @type void
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

/** Saves active Document to MarkLogic from client system.  .docx being saved to ML must already exist (be saved) on client and have both path and name.
 * @param filename the name of the file (including path) to be saved to MarkLogic  
 * @param url the url on MarkLogic that the client calls to upload the presentation
 * @param user username for MarkLogic Server url connects with
 * @param pwd password for MarkLogic Server url connects with
 * @type String
 * @throws Exception if unable to save the active Presentation
 */
MLA.saveActiveDocument = function(filename, url, user, pwd)
{
	var msg=window.external.saveActiveDocument(filename, url, user, pwd);
	var errMsg = MLA.errorCheck(msg);

        if(errMsg!=null) 
        	throw("Error: Not able to saveActiveDocument; "+errMsg);

	return msg;
}

/** Locks Content Control in active document identified by id.  Locking a Content Control disables the ability to remove the control from the document.  
 * @param id the id identifier of the Content Control to be locked in the active document.
 * @return void
 * @type void
 * @throws Exception if unable to lock Content Control
 */
MLA.lockContentControl = function(id)
{
	var msg=window.external.lockContentControl(id);
	var errMsg = MLA.errorCheck(msg);

        if(errMsg!=null) 
        	throw("Error: Not able to lockContentControl(); "+errMsg);

	if(msg=="")
	  msg = null;

	return msg;
}

/** Unlocks Content Control in active document identified by id.  Unlocking a Content Control enables the ability to remove the control from the document.  
 * @param id the id identifier of the Content Control to be unlocked in the active document.
 * @return void
 * @type void
 * @throws Exception if unable to unlock Content Control
 */
MLA.unlockContentControl = function(id)
{
	var msg=window.external.unlockContentControl(id);
	var errMsg = MLA.errorCheck(msg);

        if(errMsg!=null) 
        	throw("Error: Not able to unlockContentControl(); "+errMsg);

	if(msg=="")
	  msg = null;

	return msg;
}

/** Locks contents of Content Control in active document identified by id.  Locking Content Control contents disables editing of contents within the control.
 * @param id the id identifier of the Content Control containing contents to be locked.
 * @return void
 * @type void
 * @throws Exception if unable to lock Content Control contents
 */
MLA.lockContentControlContents = function(id)
{
	var msg=window.external.lockContentControlContents(id);
	var errMsg = MLA.errorCheck(msg);

        if(errMsg!=null) 
        	throw("Error: Not able to lockContentControlContents(); "+errMsg);

        if(msg=="")
	  msg = null;

	return msg;
}

/** Unlocks contents of Content Control in active document identified by id.  Unlocking Content Control contents enables editing of contents within the control.
 * @param id the id identifier of the Content Control containing contents to be unlocked.
 * @return void
 * @type void
 * @throws Exception if unable to unlock Content Control contents
 */
MLA.unlockContentControlContents = function(id)
{
	var msg=window.external.unlockContentControlContents(id);
	var errMsg = MLA.errorCheck(msg);

        if(errMsg!=null) 
        	throw("Error: Not able to unlockContentControlContents(); "+errMsg);

	if(msg=="")
	  msg = null;

	return msg;
}

/** Adds Content Control to active document.  Returns the id of the control added.
 * @param tag the tag identifier for the control.
 * @param title the title identifier for the control.
 * @param type the type of control.
 * @param insertPara [optional] true/false; true: inserts paragraph after control, (default)false: inserts space after control.
 * @param parentID [optional] the ID for the parent which the control should be embedded within.
 * @return id the id for the Content Control added.
 * @type String
 * @throws Exception if unable to add content control
 */

MLA.addContentControl = function(tag, title, type, insertPara, parentID)
{      //returns id of added control

       if(parentID == null || parentID == "")
       {
	       parentID="";
       }

       if(insertPara == null || insertPara == "")
       {
	       insertPara="";
       }

       var msg=window.external.addContentControl(tag,title,type,insertPara,parentID);
       var errMsg = MLA.errorCheck(msg);

       if(errMsg!=null) 
        	throw("Error: Not able to addContentControl(); "+errMsg);

       //alert("message"+msg);
       return msg;
}

/** Removes Content Control from active document.
 * @param id the id of the control to be removed.
 * @param deletecontents [optional] true/false ; true: deletes contents of control.  (default)false: removes control but document retains control contents.
 * @return void
 * @type void
 * @throws Exception if unable to remove content control
 */
MLA.removeContentControl = function(id, deletecontents)
{
       var msg=window.external.removeContentControl(id,deletecontents);
       var errMsg = MLA.errorCheck(msg);

       if(errMsg!=null) 
        	throw("Error: Not able to removeContentControl(); "+errMsg);

       return msg;
}

/** Maps Content Control in active document identified by id to Custom XML Part in .docx package.  XPath identifies which value to be displayed in active document from custom part.  Changes in custom part are immediately reflected in the active document.  Likewise, edits made to mapped control contents in the active document are immediately reflected in the custom part.
 * @param id the id identifier of the Content Control to be mapped.
 * @param xpath the xpath expression identifying which value to use for the control contents. 
 * @param prefix namespace identifier for the custom XML part.
 * @param partid the identifier of the custom XML Part which the control is to be mapped to.
 * @return void
 * @type void
 * @throws Exception if unable to map Content Control
 */
MLA.mapContentControl = function(id, xpath, prefix, partid)
{
       var msg=window.external.mapContentControl(id, xpath, prefix, partid);
       var errMsg = MLA.errorCheck(msg);

       if(errMsg!=null) 
        	throw("Error: Not able to mapContentControl(); "+errMsg);

       if(msg=="")
	  msg = null;

       return msg;
}

/** Insert text into plain or rich text control identified by id.  Text uses insertAfter() for append.  To remove control text, see the function setContentControlPlaceholderText().
 * @param id the id of the control where text will be inserted.
 * @param text the text to be inserted as the control contents.  insert is asynchronous within the active document.
 * @return void
 * @type void
 * @throws Exception if unable to insert content control text
 */
MLA.insertContentControlText = function(id, text)
{
       var msg = window.external.insertContentControlText(id,text);
       var errMsg = MLA.errorCheck(msg);

       if(errMsg!=null) 
        	throw("Error: Not able to insertContentControlText(); "+errMsg);

       return msg;
}

/** Insert image into picture control identified by id.
 * @param id the id of the control where image will be inserted.
 * @param uri the uri of the image within MarkLogic Server.
 * @param user the MarkLogic Server user.
 * @param pwd the MarkLogic Server pwd.
 * @return void
 * @type void
 * @throws Exception if unable to insert content control image
 */
MLA.insertContentControlImage = function(id, picuri, user, pwd)
{
      var msg = window.external.insertContentControlImage(id, picuri,user, pwd);
      var errMsg = MLA.errorCheck(msg);

      if(errMsg!=null) 
        	throw("Error: Not able to insertContentControlImage(); "+errMsg);

      return msg;
}

//void (similar to insertImage/Text - in C# : ContentControlListEntry

/** Adds dropdown list entries to a combobox or dropdownlist content control.  Both look like similar controls.  But combobox controls allow users to edit/add entries to the drop down list, while dropdownlist controls will only allow the user to select and use the entries available in the list.
 * @param id the id of the control to add list entries to.
 * @param text the text to be displayed in the list entry of the control.
 * @param value the value for the list entry in the control. (not displayed)
 * @param index the index where to insert the entry in the drop down list
 * @return void
 * @type void
 * @throws Exception if unable to add dropdown list entries.
 */
MLA.addContentControlDropDownListEntries = function(id, text, value, index)
{
      var msg = window.external.addContentControlDropDownListEntries(id, text, value, index);
      var errMsg = MLA.errorCheck(msg);

      if(errMsg!=null) 
        	throw("Error: Not able to addContentControlDropDownListEntries(); "+errMsg);

      return msg;
}

/**
 * Returns selected content control dropdown list entry range text specified by id in the active document.
 * @returns the text for the selected dropdown list entry. If nothing is selected the placeholder text (if available) is returned.  If there is no placeholder text, then the empty string is returned.
 * @type String 
 * @throws Exception if unable to get content control dropdown list entry selected text.
 */
MLA.getContentControlDropDownListEntrySelectedText = function(id)
{
      var msg = window.external.getContentControlDropDownListEntrySelectedText(id);
      var errMsg = MLA.errorCheck(msg);

      if(errMsg!=null) 
        	throw("Error: Not able to getContentControlDropDownListEntrySelectedText(); "+errMsg);

      return msg;
}

 /**
 * Returns selected content control dropdown list entry value specified by id in the active document.
 * @returns the value for the selected dropdown list entry. If nothing is selected, then the empty string is returned.
 * @type String 
 * @throws Exception if unable to get content control dropdown list entry selected value.
 */
MLA.getContentControlDropDownListEntrySelectedValue = function(id)
{
      var msg = window.external.getContentControlDropDownListEntrySelectedValue(id);
      var errMsg = MLA.errorCheck(msg);

      if(errMsg!=null) 
        	throw("Error: Not able to getContentControlDropDownListEntrySelectedValue(); "+errMsg);

      return msg;
}

/**
 * Returns ids for all content controls in the active document.
 * @returns the ids for all content controls in the active document as array of string
 * @type Array 
 * @throws Exception if unable to get content control ids.
 */
MLA.getContentControlIds = function()
{
        var controlIds = window.external.getContentControlIds();
	var errMsg = MLA.errorCheck(controlIds);

	if(errMsg!=null)
	  throw("Error: Not able to getContentControlIds(); "+errMsg);

	var contentControlIds = controlIds.split("|");

	if(contentControlIds.length ==1)
	{
		if (contentControlIds[0] == null || contentControlIds[0] == "")
		{
			contentControlIds.length = 0;
		}
	}

	return contentControlIds;

}

/** @ignore */
MLA.getContentControlInfo = function(cid)
{ 
	var info = window.external.getContentControlInfo(cid);
	var errMsg = MLA.errorCheck(info);

	if(errMsg!=null)
	  throw("Error: Not able to getContentControlInfo(); "+errMsg);

	//alert("info: "+info);
	return info;

}

//SimpleContentControl
/**
 * Returns SimpleContentControl for nearest parent control of current cursor position in the active document. 
 * @returns a SimpleContentControl for the parent content control of the nearest control that the cursor is embedde within in the active document.
 * @type SimpleContentControl 
 * @throws Exception if unable to get parent content control info.
 */
MLA.getParentContentControlInfo = function()
{
        var info = window.external.getParentContentControlInfo();
	var errMsg = MLA.errorCheck(info);

        if(errMsg!=null)
	  throw("Error: Not able to getParentContentControlInfo(); "+errMsg);


	var tokens = info.split("|");

	var controlid = tokens[0];
        var mlacontrolref = new MLA.SimpleContentControl(controlid); 
	    mlacontrolref.tag = tokens[1];
            mlacontrolref.title = tokens[2];  
	    mlacontrolref.type = tokens[3];
            mlacontrolref.parentTag = tokens[4];
            mlacontrolref.parentID = tokens[5];

	return mlacontrolref;
}

/**
 * Sets focus in active document to beginning of control specified by id. 
 * @param id the id identifying the control to which focus will be set in active document.
 * @return void
 * @type void
 * @throws Exception if unable to set content control focus.
 */
MLA.setContentControlFocus= function(id)
{
   var msg = window.external.setContentControlFocus(id);

   var errMsg = MLA.errorCheck(msg);

   if(errMsg!=null)
      throw("Error: Not able to setContentControlFocus(); "+errMsg);

   return msg;

}

/** Insert image into picture control identified by id.
 * @param id the ID of the control in which to set placeholder text.
 * @param placeholdertext the placeholder text to be used for the control.
 * @param cleartext [optional] true/false; true: delete existing text for control range. (default)false: do not delete control range text contents.
 * @return void
 * @type void
 * @throws Exception if unable to set content control placeholder text
 */
MLA.setContentControlPlaceholderText= function(id, placeholdertext, cleartext)
{

    if (cleartext == null || cleartext == "")
    {
       cleartext="false";
    }

    var msg = window.external.setContentControlPlaceholderText(id, placeholdertext, cleartext);
    var errMsg = MLA.errorCheck(msg);


    if(errMsg!=null)
      throw("Error: Not able to setContentControlPlaceholderText(); "+errMsg);

    return msg;
       
}

/**
 * Returns text for range of content control identified by id in active document. 
 * @id the id of the control
 * @returns text contents of range within control
 * @type String 
 * @throws Exception if unable to get content control text.
 */
MLA.getContentControlText = function(id)
{
    var msg = window.external.getContentControlText(id);
    var errMsg = MLA.errorCheck(msg);

    if(errMsg!=null)
      throw("Error: Not able to getContentControlText(); "+errMsg);

    return msg;
       
}

/** Returns WordOpenXML property for range of control in active document specified by id.  The WordOpenXML property returns the control as if it were its own Word document, saved as .xml.  The representation is also called Flat OPC, following the Open Packaging Convention.
 *
 * @param id the id of the control.
 * @return Microsoft.XMLDOM object that is the content control contents in the WordOpenXML property format.
 * @type Microsoft.XMLDOM object
 * @throws Exception if unable to get content control WordOpenXML
 */
MLA.getContentControlWordOpenXML = function(id)
{
    var msg = window.external.getContentControlWordOpenXML(id);
    var errMsg = MLA.errorCheck(msg);

    if(errMsg!=null)
      throw("Error: Not able to getContentControlWordOpenXML(); "+errMsg);

    var v_documentXml = MLA.createXMLDOM(msg);
    return v_documentXml;
       
}

/** Returns WordOpenXML property for selection in active document.  If nothing is selected, the document.xml part will have an empty body. The WordOpenXML property returns the selection as if it were its own Word document, saved as .xml.  The representation is also called Flat OPC, following the Open Packaging Convention.
 *
 * @return Microsoft.XMLDOM object that is the selection in the WordOpenXML property format.
 * @type Microsoft.XMLDOM object
 * @throws Exception if unable to get selection WordOpenXML
 */
MLA.getSelectionWordOpenXML = function()
{
    var v_documentXml;
    var msg = window.external.getSelectionWordOpenXML();
    var errMsg = MLA.errorCheck(msg);

    if(errMsg!=null)
      throw("Error: Not able to getSelectionWordOpenXML(); "+errMsg);

    if(msg=="")
    {
	v_documentXMl=null; 
    }
    else
    {
	v_documentXml = MLA.createXMLDOM(msg);
    }

    return v_documentXml;
}

/** Returns WordOpenXML property for the active document.  The WordOpenXML property returns the document as a Word document saved as .xml.  This representation is also called Flat OPC, following the Open Packaging Convention.
 *
 * @return Microsoft.XMLDOM object that is the active document in the WordOpenXML property format.
 * @type Microsoft.XMLDOM object
 * @throws Exception if unable to get the document WordOpenXML
 */
MLA.getDocumentWordOpenXML = function()
{
    var v_documentXml;
    var msg = window.external.getDocumentWordOpenXML();
    var errMsg = MLA.errorCheck(msg);

    if(errMsg!=null)
      throw("Error: Not able to getDocumentWordOpenXML(); "+errMsg);

    if(msg=="")
    {
	v_documentXMl=null; 
    }
    else
    {
	v_documentXml = MLA.createXMLDOM(msg);
    }

    return v_documentXml;
}
/** Sets WordOpenXML for the active document.  The WordOpenXML property is read only in the Word Object Model.  This function however rewrites the active document package with the XML passed here as a parameter.  The package representation is also called Flat OPC, following the Open Packaging Convention.
 *
 * @param opc_xml the XML to be inserted. Parameter type can be either A) an XMLDOM object that is the WordOpenXML to be inserted into the active Open XML package, or B)the string serialization of the WordOpenXML to be inserted into the active Open XML package
 * @return void
 * @type void
 * @throws Exception if unable to setDocumentWordOpenXML
 */
MLA.setDocumentWordOpenXML = function(opc_xml)
{
     var v_docx="";

     if(opc_xml.xml)
     { 
        v_docx = opc_xml.xml;
     }
     else
     { 
	v_docx = opc_xml;
     }

    var msg = window.external.setDocumentWordOpenXML(v_docx);
    var errMsg = MLA.errorCheck(msg);

    if(errMsg!=null)
      throw("Error: Not able to setDocumentWordOpenXML(); "+errMsg);

    return msg;
}

/** Sets the tag for content control specified by id in active document 
 *
 * @param id the id of the control.
 * @param tag the tag to be set for the control.
 * @return void
 * @type void
 * @throws Exception if unable to set content control tag
 */
MLA.setContentControlTag = function(id, tag)
{
    var msg = window.external.setContentControlTag(id,tag);
    var errMsg = MLA.errorCheck(msg);

    if(errMsg!=null)
      throw("Error: Not able to setContentControlTag(); "+errMsg);

    return msg;
       
}

/** Sets the title for content control specified by id in active document 
 *
 * @param id the id of the control.
 * @param title the title to be set for the control.
 * @return void
 * @type void
 * @throws Exception if unable to set content control title
 */
MLA.setContentControlTitle = function(id, title)
{
    var msg = window.external.setContentControlTitle(id, title);
    var errMsg = MLA.errorCheck(msg);

    if(errMsg!=null)
      throw("Error: Not able to setContentControlTitle(); "+errMsg);

    return msg;
}

/** Sets the style for the content control specified by id in active document 
 *
 * @param id the id of the control.
 * @param style the style property to be set for the control.  This style can be any of the styles available in Office.  For example "Subtitle","Heading 1","Heading 2", etc.  For more options, see the 'styles' group on the 'Home' tab of Word.
 * @return void
 * @type void
 * @throws Exception if unable to set content control style
 */
MLA.setContentControlStyle = function(id, style)
{
    var msg = window.external.setContentControlStyle(id, style);
    var errMsg = MLA.errorCheck(msg);

    if(errMsg!=null)
      throw("Error: Not able to setContentControlStyle(); "+errMsg);

    return msg;
}

/**
 * Returns array of SimpleContentControl objects. A SimpleContentControl is returned for each content control in the active document. Properties include:
 *
 * @returns array of SimpleContentControl objects.
 * @type Array 
 * @throws Exception if unable to get SimpleContentControls.
 */
MLA.getSimpleContentControls = function()
{
        var controlIds = window.external.getContentControlIds();
	var errMsg = MLA.errorCheck(controlIds);

	if(errMsg!=null)
	  throw("Error: Not able to getContentControlIds(); "+errMsg);

	var contentControlIds = controlIds.split("|");
        var controlArray = new Array();
      

	if (contentControlIds[0] == null || contentControlIds[0] == "")
	{
			contentControlIds.length = 0;
		
	}else
	{
	       for(var i =0; i<contentControlIds.length; i++)
	       {
	  	  var controlid = contentControlIds[i];
	          var mlacontrolref = new MLA.SimpleContentControl(controlid);
		  var contentControlInfo = MLA.getContentControlInfo(controlid);

		  var info = contentControlInfo.split("|");

		  mlacontrolref.tag = info[0];
		  mlacontrolref.title = info[1]; 
		  mlacontrolref.type = info[2];
		  mlacontrolref.parentTag = info[3]; 
		  mlacontrolref.parentID = info[4]; 

		  controlArray[i]=mlacontrolref;
	       }
	}

	return controlArray;
}

/** Inserts WordOpenXML at the current cursor position in the active document.  WordOpenXML is the XML for a Word document in Flat OPC format (the Open Packaging Convention).  It is the same format as if you'd saved a .docx as .xml in Word.  
 *
 * @param opc_xml the XML to be inserted. Parameter type can be either A) an XMLDOM object that is the WordOpenXML to be inserted into the active Open XML package, or B)the string serialization of the WordOpenXML to be inserted into the active Open XML package.
 * @return void
 * @type void
 * @throws Exception if unable to insertWordOpenXML
 */
MLA.insertWordOpenXML = function(opc_xml)
{
     var v_docx="";

     if(opc_xml.xml)
     { 
        v_docx = opc_xml.xml;
     }
     else
     { 
	v_docx = opc_xml;
     }

    var msg = window.external.insertWordOpenXML(v_docx);
    var errMsg = MLA.errorCheck(msg);

    if(errMsg!=null)
      throw("Error: Not able to insertWordOpenXML(); "+errMsg);

    return msg;
}

/** @ignore */
MLA.hideContentControlRange = function(id)
{
    var msg = window.external.hideContentControlRange(id);
    var errMsg = MLA.errorCheck(msg);

    if(errMsg!=null)
      throw("Error: Not able to insertWordOpenXML(); "+errMsg);

    return msg;
}
/** @ignore */
MLA.displayContentControlRange = function(id)
{
    var msg = window.external.displayContentControlRange(id);
    var errMsg = MLA.errorCheck(msg);

    if(errMsg!=null)
      throw("Error: Not able to insertWordOpenXML(); "+errMsg);

    return msg;
}

/**
 * Create a new SimpleContentControl instance. 
 * @class A basic SimpleContentControl class.
 */
MLA.SimpleContentControl = function()
{
 
  var title;
  var tag;
  var id;
  var type;
  var parentTag;
  var parentID;

  switch (typeof arguments[0])
  {
    //case 'number' : MLA.ContentControlRef.$int.apply(this, arguments); break;
    case 'string' : MLA.SimpleContentControl.$str.apply(this, arguments); break;
    default : /*NOP*/
  }
  

}
/**
 * Create a new SimpleContentControl instance. 
 * @class a basic SimpleContentControl class.
 * @constructor
 * @param {string} the ID property of the ContentControl
 * @see MLA.SimpleContentControl() is the base class for this
 */
MLA.SimpleContentControl.$str = function(cid) {
this.id= cid;
//var cell = MLA.convertA1ToR1C1(coord);
//var c_values = cell.split(":");
//this.rowIdx= c_values[1];
//this.colIdx=c_values[0];
//
}

