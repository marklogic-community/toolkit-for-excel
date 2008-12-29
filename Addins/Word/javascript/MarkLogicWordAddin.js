/* 
Copyright 2008 Mark Logic Corporation

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
 * @version 0.1 
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
/** Utility function to create a default WordprocessingML paragraph <w:p>, with no styles, for a given string.
 *
 *@param textstring the string to be converted into a WordprocessingML paragraph
 *@return Microsoft.XMLDOM object that is a WordprocessingML paragraph
 *@throws Exception if unable to create the paragraph
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
 *@param textToInsert the text to be inserted at the current cursor position in the ActiveDocument.
 *@throws Exception if unable to insert text 
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
 *@return the text selected by the user in the ActiveDocument as strings. If nothing is selected, an empty array is returned.
 *@type Array
 *@param delimiter (optional) Text from tables will be returned as a single string, with cells delimited by tabs (default).  This param may be used to assign a different delimiter.  Note: there is no delimiter for text from paragraphs, as each paragraph will be captured in separate array elements.
 *@throws Exception if unable to retrieve the Selection
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

/** Returns the XML that represents what is currently selected (highlighted) by the user in the ActiveDocument as an XMLDOM object.  Whatever is highlighted by the user will be returned in this function as a block level element.  A user may highlight text that will be materialized as multiple sibling block elements in the XML.  For this reason, the function returns an array, where each element of the array is an XMLDOM object that contains the XML for the blocks highlighted by the user in the ActiveDocument.  The order of elements in the array represents the order of items that are highlighted in the ActiveDocument.
 *@return the blocks of XML currently selected by the user in the ActiveDocument as XMLDOM objects. If nothing is selected, an empty array is returned.
 *@type Array
 *@throws Exception if unable to retrieve the Selection
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
 *@return the block of XML, as XMLDOM object,  for the Sentence immediately preceding the cursor position.
 *@type Microsoft.XMLDOM object
 *@throws Exception if unable to retrieve the XML for the Sentence.
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
 *@throws Exception if unable to retrieve Styles.xml from the ActiveDocument.
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

/** Inserts block-level xml into the ActiveDocument in Word at the current cursor position, or over the selected range (if selected).  If stylesxml parameter is defined, the blockxml will be inserted and the Styles.xml will be overwritten with the stylesxml contents. If stylesxml is not defined, then only the blockxml will be inserted ; the assumption being that whatever styles required by the blockxml are already present in the .docx Styles.xml.
 *
 * There are two main levels of content in the document.xml; block-level and inline content. Block level describes the structure of the document, and includes paragraphs and tables. Anything that can be inserted under <w:body> may be inserted here
 * 
 *@param blockXml this parameter may either be A) and XMLDOM object that is the block-level XML to be inserted,or an array of such objects,or B) a String, that is the serialized, well-formed XML of the block to be inserted, or an Array of such Strings.
 *@param stylesXml (optional) this parameter is either A) an XMLDOM object that contains Styles.xml for the pacakge, or B)a String, which is the serialized, well-formed XML that represents Styles.xml for the ActiveDocument package in Word.
 *@throws Exception if unable to insert the blockContent or Styles.xml
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
