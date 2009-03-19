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

MarkLogicExcelAddin.js - javascript api for interacting with webBrowser control within Custom Task Pane enabled in Excel.
*/

/** 
 * @fileoverview  API documentation for MarkLogicExcelAddin.js
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
 * The functions here provide ways for interacting with the  active Workbooks, Worksheets, and Cells in Excel; however, the functions getCustomXMLPart(), getCustomXMLPartIds(), addCustomXMLPart(), and deleteCustomXMLPart() will work for any Open XML package, provided they are used within the context of an Addin for the appropriate Office application.
 */
var MLA = {};
/*
function MLA(){
      this.getClassName = function(){
      return "MLA";
   
      }
}

*/

/**
 * Create a new Cell instance. 
 * @class A basic Cell class.
 * @constructor
 */

MLA.Cell = function(){
 
  var rowIdx;
  var colIdx;
  var coordinate;
  var value2;
  var formula;

  switch (typeof arguments[0])
  {
    case 'number' : MLA.Cell.$int.apply(this, arguments); break;
    case 'string' : MLA.Cell.$str.apply(this, arguments); break;
    default : /*NOP*/
  }
/* 
  this.setFormula = function(f)
  {
	  this.formula=f;
  }

  this.setValue2 = function(v)
  {
	  this.value2=v;
  }

  this.getFormula = function()
  {
	  return this.formula;
  }

  this.getValue2 = function()
  {
	  return this.value2;
  }

*/
  //alert("In the constructor"+this.rowIdx+" "+this.colIdx+" "+this.coordinate+" "+this.value2);
	
}

/**
 * Create a new Cell instance. 
 * @class A basic Cell class.
 * @constructor
 * @param {int} the x coordinate (R1 value) on the spreadsheet
 * @param {int} the y coordinate (C1 value) on the spreadsheet
 * @see MLA.Cell() is the base class for this
 */
MLA.Cell.$int = function(x,y) {
this.rowIdx = x;
this.colIdx = y;
this.coordinate=MLA.convertR1C1ToA1(x,y);
};

MLA.Cell.$str = function(coord) {
this.coordinate = coord;
var cell = MLA.convertA1ToR1C1(coord);
var c_values = cell.split(":");
this.rowIdx= c_values[1];
this.colIdx=c_values[0];
}

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
 * Returns version of MarkLogicExcelAddin.js library
 * @return the version of MarkLogicExcelAddin.js
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

/** @ignore */
MLA.isArray = function(obj)
{
 return obj.constructor == Array;
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

//START EXCEL ONLY FUNCTIONS
//NEED TO ADD ERROR HANDLING, HERE AND IN THE ADDIN (C#)

/**
 * Returns the name of the active workbook.
 * @return name of active workbook
 * @type String 
 * @throws Exception if unable to retrieve the active workbook name 
 */
MLA.getActiveWorkbookName = function()
{
   	var wbName = window.external.getActiveWorkbookName();

   	var errMsg = MLA.errorCheck(wbName);

   	if(errMsg!=null)
        	throw("Error: Not able to getActiveWorkbookName; "+errMsg);

//simple template for error checking
//	var errMsg = MLA.errorCheck(newId);
//	if(errMsg!=null)
//	   throw("Error: Not able to addCustomXMLPart; "+errMsg);
//
//	if(newId =="")
//	  newId=null

   	return wbName;
}

/**
 * Returns the name of the active worksheet.
 * @return name of active worksheet
 * @type String 
 * @throws Exception if unable to retrieve the active worksheet name
 */
MLA.getActiveWorksheetName = function()
{
	var wsName = window.external.getActiveWorksheetName();

	var errMsg = MLA.errorCheck(wsName);

        if(errMsg!=null)
        	throw("Error: Not able to getActiveWorksheetName; "+errMsg);

	return wsName;
}

/**
 * Returns the names of all workbooks available as array of strings.
 * @return names of all workbooks
 * @type Array 
 * @throws Exception if unable to retrieve the workbook names
 */
MLA.getAllWorkbookNames = function()
{
	var names = window.external.getAllWorkbookNames();
	var errMsg = MLA.errorCheck(names);
	
        if(errMsg!=null)
        	throw("Error: Not able to getAllWorkbookNames; "+errMsg);

	var wbnames = names.split("|");  //change this to JSON

        if(wbnames.length ==1)
	{
		if (wbnames[0] == null || wbnames[0] == "")
		{
			wbnames.length = 0;
		}
	}

	return wbnames;

}


/**
 * Returns the names of all worksheets for active workbook as array of strings.
 * @return names of worksheets for active workbook
 * @type Array 
 * @throws Exception if unable to retrieve the worksheet names
 */
MLA.getActiveWorkbookWorksheetNames = function()
{
	var names = window.external.getActiveWorkbookWorksheetNames();

	var errMsg = MLA.errorCheck(names);
	
        if(errMsg!=null)
        	throw("Error: Not able to getActiveWorkbookWorksheetNames; "+errMsg);

	var wsnames = names.split("|");  //change this to JSON

        if(wsnames.length ==1)
	{
		if (wsnames[0] == null || wsnames[0] == "")
		{
			wsnames.length = 0;
		}
	}

	return wsnames;

}

/**
 * Adds workbook of type Excel.XlWBATemplate.xlWBATWorksheet.  Added workbook receives focus.
 * @param worksheetname - a string that names the worksheet 
 * @return name of workbook added - Note: the only way to name the workbook is to Save it.  Excel returns a default name for the workbook, but this name will change the first time the user saves the workbook. 
 * @type String 
 * @throws Exception if unable to add the workbook
 */
MLA.addWorkbook = function(worksheetname) //,subject,saveas)
{
	var wb = window.external.addWorkbook(worksheetname); //,subject,saveas);
	var errMsg = MLA.errorCheck(wb);
	
        if(errMsg!=null)
        	throw("Error: Not able to addWorkbook; "+errMsg);

	return wb;
}

/**
 * Adds worksheet to active workbook.  Worksheet will be added to end of existing sheets in workbook at the final position. Added worksheet receives focus.
 * @param name - the name of the worksheet to be added
 * @throws Exception if unable to add the worksheet
 */
MLA.addWorksheet = function(name)
{
	var ws = window.external.addWorksheet(name);

	var errMsg = MLA.errorCheck(ws);
	
        if(errMsg!=null)
        	throw("Error: Not able to addWorksheet; "+errMsg);

	return ws;
}

/**
 * Sets the active workbook in Excel. 
 * @param workbookname - a string that is the name of the workbook to receive focus 
 * @throws Exception if unable to activate the specified workbook
 */
MLA.setActiveWorkbook = function(workbookname)
{
	var saw = window.external.setActiveWorkbook(workbookname);
	var errMsg = MLA.errorCheck(saw);
	
        if(errMsg!=null)
        	throw("Error: Not able to setActiveWorkbook; "+errMsg);

	//return saw;
}

/**
 * Sets the active worksheet in Excel.   
 * @param worksheetname - a string that is the name of the worksheet to receive focus 
 * @throws Exception if unable to activate the specified worksheet
 */
MLA.setActiveWorksheet = function(sheetname)
{
	var saw = window.external.setActiveWorksheet(sheetname);
	var errMsg = MLA.errorCheck(saw);
	
        if(errMsg!=null)
        	throw("Error: Not able to setActiveWorksheet; "+errMsg);

	//return saw;
}


MLA.addNamedRange = function(coord1,coord2,rngName)
{
	var nr = window.external.addNamedRange(coord1,coord2,rngName);
	return nr;
}

MLA.addAutoFilter = function(coord1, coord2, criteria1, operator, criteria2)
{

	if(criteria1==null)
	{
		criteria1="<>";
	}

	if(criteria2==null)
	{
		criteria2="missing";
	}

	if(operator==null)
	{
		operator="AND";
	}

	var rng = window.external.addAutoFilter(coord1,coord2,criteria1,operator,criteria2);
	return rng;
}

MLA.getNamedRangeRangeNames = function()
{
	var nrs = window.external.getNamedRangeRangeNames();
	var nrsArray = nrs.split(":");
	return nrsArray;
}

MLA.setActiveRangeByName = function(name)
{
	var msg = window.external.setActiveRangeByName(name);
	return msg;
}

MLA.clearNamedRange = function(name)
{
	var msg=window.external.clearNamedRange(name);
	return msg;
}

MLA.clearRange = function(scoord,ecoord)
{
	var msg=window.external.clearRange(scoord,ecoord);
	return msg;
}

MLA.removeNamedRange = function(name)
{
	var msg = window.external.removeNamedRange(name);
	return msg;
}

MLA.getSelectedRangeCoordinates = function()
{
        var r = window.external.getSelectedRangeCoordinates();
	return r;
}

MLA.getSelectedCells = function()
{
     var cellresults = window.external.getSelectedCells();
     var cellstring = "{ \"cells\" : "+cellresults+"}";
     //alert (cellstring);
     var cells = eval('('+cellstring+')');
     //alert(cells.cells.length);
     var cellArray = new Array();
     for(var i =0;i<cells.cells.length;i++)
     {
	     var cell = cells.cells[i];
	     var mlacell = new MLA.Cell();
             mlacell=cell;
	     cellArray[i]=mlacell;
	     //alert("coordinate: "+cell.coordinate+"arraylength"+cellArray.length);
	     //alert("cell Coordinate: "+cellArray[i].coordinate+" rowIdx: "+cellArray[i].rowIdx+" colIdx: "+cellArray[i].colIdx+" value2: "+cellArray[i].value2);
     }

     return cellArray;
}

MLA.getActiveCell = function()
{
	var cellinfo = window.external.getActiveCell();
	var cellValues = cellinfo.split(":");
	var rowIdx = cellValues[0];
	var colIdx = cellValues[1];
	var newCell = new MLA.Cell(parseInt(rowIdx),parseInt(colIdx));

	if(cellValues[2]=="")
	{
		newCell.value2=null;
	}else
	{
	        newCell.value2 = cellValues[2];
	}

	if(cellValues[3]=="")
	{
		newCell.formula=null;
	}else
	{
		newCell.formula = cellValues[3];
	}

	return newCell;
}

/** @ignore */
MLA.getActiveCellRange = function()
{
	var cell = window.external.getActiveCellRange();
	return cell;
}

/**
 * Get the Text for the cell at the current cursor position.
 * @return the text in the currently selected cell
 * @type String
 * @throws Exception if unable to return the active cell text
 */
MLA.getActiveCellText = function()
{
	var cellstring = window.external.getActiveCellText();
	var cell = eval('('+cellstring+')');
	var mlaCell = new MLA.Cell();
	mlaCell = cell;
	return mlaCell;
}

/**
 * Set the Text for the cell at the current cursor position.
 * @param value - the text to be inserted in the active cell
 * @throws Exception if unable to set the active cell text
 */
MLA.setActiveCellValue = function(value)
{
	var msg = window.external.setActiveCellValue(value);
        return msg;
}


MLA.setCellValue = function(cells)
{ 
	//alert("IN FUNCTION");

	var v_array = MLA.isArray(cells);

	if(v_array)
	{
		for(var i =0; i<cells.length; i++)
		{
          		var msg = window.external.setCellValueA1(cells[i].coordinate, cells[i].value2);
		}
	}

	return msg;
}

MLA.convertA1ToR1C1 = function(coord)
{
	var msg=window.external.convertA1ToR1C1(coord);
	return msg;
}

MLA.convertR1C1ToA1 = function(rowIdx, colIdx)
{
	var msg=window.external.convertR1C1ToA1(rowIdx, colIdx);
	return msg;
}

MLA.clearWorksheet = function()
{
	var msg=window.external.clearActiveWorksheet();
        return msg;
}

MLA.getTempPath = function()
{
	//alert("IN HERE");
	var msg=window.external.getTempPath();
	return msg;
}


/*
MLA.setCellValueR1C1 = function(cells)
{
	alert("IN R1C1 FUNCTION");

	var v_array = MLA.isArray(cells);

	if(v_array)
	{
		for(var i =0; i<cells.length; i++)
		{
          		var msg = window.external.setCellValueA1(cells[i].rowIdx, cells[i].colIdx, cells[i].value2);
		}
	}

	return msg;
}
*/
