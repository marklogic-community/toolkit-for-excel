/* 
Copyright 2009-2011 MarkLogic Corporation

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
 * @version 1.0-2 
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
 */
MLA.Cell = function(){
 
  var rowIdx;
  var colIdx;
  var coordinate;
  var value2;
  var formula;

  switch (typeof arguments[0]){
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
MLA.version = { "release" : "2.0-0" }; 

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
 * Returns version of MarkLogicExcelAddin.js library.
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

/** Utility function for creating Microsoft.XMLDOM object from string.
 *
 *@param xmlString the string to be loaded into a XMLDOM object.  The string must be serialized, well-formed XML
 *@return Microsoft.XMLDOM object
 *@throws Exception if unable to create the XMLDOM object
 */
MLA.createXMLDOM = function(xmlstring)
{
	try{
		var xmlDom = new ActiveXObject("Microsoft.XMLDOM");
       		xmlDom.async=false;
       		xmlDom.loadXML(xmlstring);
	}catch(err){
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

	if(customPartIds.length ==1){
		if (customPartIds[0] == null || customPartIds[0] == ""){
			customPartIds.length = 0;
		}
	}

	return customPartIds;
}

/**
 * Returns the custom XML part, identified by customXMLPartId, that is part of the active Open XML package. (.docx, .xlsx, .pptx, etc.)
 * @param customXMLPartId the id of the custom part to be fetched from the active package
 * @return the XML for the custom part as a DOM object
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

	if(customXMLPart==""){
		v_cp=null;
	}else{
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
	if(customPartXml.xml){
               v_customPart=customPartXml.xml;
	}
	else{
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
 *@param customXMLPartId the id of the custom part to be deleted from the active Open XML package
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
version - the version of the Addin library, url - the url used by the Addin WebBrowser control, theme - the current color scheme used by Office
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
 * Returns the type of the sheet specified by sheetName.
 * @param sheetName the name of the sheet to inspect
 * @return type of sheet
 * @type String 
 * @throws Exception if unable to get sheet type
 */
MLA.getSheetType = function(sheetName)
{
        var wsType = window.external.getSheetType(sheetName);

	var errMsg = MLA.errorCheck(wsType);

        if(errMsg!=null)
        	throw("Error: Not able to getSheetName; "+errMsg);

	return wsType;
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

        if(wbnames.length ==1){
		if (wbnames[0] == null || wbnames[0] == ""){
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

        if(wsnames.length ==1){
		if (wsnames[0] == null || wsnames[0] == ""){
			wsnames.length = 0;
		}
	}

	return wsnames;
}

/**
 * Adds workbook of type Excel.XlWBATemplate.xlWBATWorksheet.  Added workbook receives focus.
 * @param worksheetname a string that names the worksheet 
 * @return name of workbook added - Note: the only way to name the workbook is to Save it.  Excel returns a default name for the workbook, but this name will change the first time the user saves the workbook
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
 * @param sheetName the name of the worksheet to be added
 * @throws Exception if unable to add the worksheet
 */
MLA.addWorksheet = function(sheetName)
{
	var ws = window.external.addWorksheet(sheetName);

	var errMsg = MLA.errorCheck(ws);
	
        if(errMsg!=null)
        	throw("Error: Not able to addWorksheet; "+errMsg);

	return ws;
}

/**
 * Sets the active workbook in Excel. 
 * @param workbookname a string that is the name of the workbook to receive focus 
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
 * @param worksheetname a string that is the name of the worksheet to receive focus 
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

/**
 * Names a range in the worksheet specified by sheetName. 
 * @param coord1 starting coordinate of range in A1 notation
 * @param coord2 end coordinate of range in A1 notation
 * @param rngName the name to be assigned to the range 
 * @param sheetName (optional) default 'active' ; the name of the worksheet to host the range 
 * @throws Exception if unable to name the specified range
 */
MLA.addNamedRange = function(coord1,coord2,rngName, sheetName)
{

	if(sheetName==null || sheetName==""){
		sheetName="active";
	}

	var nr = window.external.addNamedRange(coord1,coord2,rngName,sheetName);
        var errMsg = MLA.errorCheck(nr);
	
        if(errMsg!=null)
        	throw("Error: Not able to name specified range; "+errMsg)
	return nr;
}

/**
 * Adds AutoFilter to specified range in active worksheet.  
 * @param coord1 starting coordinate of range in A1 notation
 * @param coord2 end coordinate of range in A1 notation
 * @param sheetName (optional) default 'active'
 * @param criteria1 (optional) default '<>'
 * @param operator (optional) default 'AND'
 * @param criteria2 (optional) default 'missing'
 * @throws Exception if unable to add AutoFilter specified range
 */
MLA.addAutoFilter = function(coord1, coord2, sheetName, criteria1, operator, criteria2)
{
        if(sheetName==null || sheetName==""){
		sheetName="active";
	}

	if(criteria1==null){
		criteria1="<>";
	}

	if(criteria2==null){
		criteria2="missing";
	}

	if(operator==null){
		operator="AND";
	}

	var rng = window.external.addAutoFilter(coord1,coord2,sheetName,criteria1,operator,criteria2);
	var errMsg = MLA.errorCheck(rng);
	
        if(errMsg!=null)
        	throw("Error: Not able to addAutoFilter; "+errMsg);

	return rng;
}

/**
 * Returns all NamedRange names for the active workbook.
 * @type Array
 * @throws Exception if unable to add retrieve NamedRange names
 */
MLA.getNamedRangeNames = function()
{
	var nrs = window.external.getNamedRangeRangeNames();

        var errMsg = MLA.errorCheck(nrs);
	
        if(errMsg!=null)
        	throw("Error: Not able to getNamedRangeRangeNames; "+errMsg);

	var nrsArray = nrs.split(":");
	return nrsArray;
}

/**
 * Returns all Chart names for any charts found on sheet specified by sheetName.
 * @param sheetName the name of the sheet containing the charts
 * @type Array
 * @throws Exception if unable to add retrieve Chart names
 */
MLA.getWorksheetChartNames = function(sheetName)
{
	var nrsArray="";
        var nrs = window.external.getWorksheetChartNames(sheetName);

        var errMsg = MLA.errorCheck(nrs);
	
        if(errMsg!=null)
        	throw("Error: Not able to getWorksheetChartNames; "+errMsg);

        if(nrs.length > 0)
	   nrsArray = nrs.split(":");

	return nrsArray;
}

/**
 * Returns all named range names for any named ranges found on sheet specified by sheetName.
 * @param sheetName the name of the sheet containing the charts
 * @type Array
 * @throws Exception if unable to add retrieve named range names
 */
MLA.getWorksheetNamedRangeNames = function(sheetName)
{
	var nrsArray="";
        var nrs = window.external.getWorksheetNamedRangeRangeNames(sheetName);

        var errMsg = MLA.errorCheck(nrs);
	
        if(errMsg!=null)
        	throw("Error: Not able to getWorksheetNamedRangeRangeNames; "+errMsg);

        if(nrs.length > 0)
	   nrsArray = nrs.split(":");

	return nrsArray;
	
}
/**
 * Returns all NamedRange names for the active workbook.
 * @param name the name of the range to be set active in the workbook 
 * @throws Exception if unable to add retrieve NamedRange names
 */
MLA.setActiveRangeByName = function(name)
{
	var msg = window.external.setActiveRangeByName(name);
	var errMsg = MLA.errorCheck(msg);
	
        if(errMsg!=null)
        	throw("Error: Not able to setActiveRangeByName; "+errMsg);

	return msg;
}

/**
 * Clears all cells in the range identified by name.
 * @param rangeName the name of the range to be cleared in the active workbook 
 * @throws Exception if unable to clear the cells in the NamedRange
 */
MLA.clearNamedRange = function(rangeName)
{
	var msg=window.external.clearNamedRange(rangeName);
	var errMsg = MLA.errorCheck(msg);
	
        if(errMsg!=null)
        	throw("Error: Not able to clearNamedRange; "+errMsg);

	return msg;
}

/**
 * Clears all cells in the range identified by coordinates provide in A1 notation.
 * @param coord1 starting coordinate of range to be cleared in A1 notation
 * @param coord2 end coordinate of range to be cleared in A1 notation 
 * @throws Exception if unable to clear the cells in the range
 */
MLA.clearRange = function(coord1,coord2)
{
	var msg=window.external.clearRange(coord1,coord2);
        var errMsg = MLA.errorCheck(msg);
	
        if(errMsg!=null)
        	throw("Error: Not able to clearRange; "+errMsg);

	return msg;
}

/**
 * Removes the NamedRange from the active workbook.  Note - cells and values stay intact, this only removes the name from the range.
 * @param name the name of the NamedRange to be removed from the active workbook
 * @param sheetName (optional) default 'active'
 * @throws Exception if unable to remove the named range
 */
MLA.removeNamedRange = function(name)
{
	var msg = window.external.removeNamedRange(name);
        var errMsg = MLA.errorCheck(msg);

	if(errMsg!=null)
        	throw("Error: Not able to removeNamedRange; "+errMsg);

	return msg;
}

/**
 * Returns the selected range coordinates.  This works for contiguous ranges.  When disparate cells are selected,  the last coordinates for the last contigous range selected in the active workbook will be returned.
 * @type String
 * @throws Exception if unable to retrieve the coordinates
 */
MLA.getSelectedRangeCoordinates = function()
{
        var msg = window.external.getSelectedRangeCoordinates();

        var errMsg = MLA.errorCheck(msg);

	if(errMsg!=null)
        	throw("Error: Not able to setSelectedRangeCoordinates; "+errMsg);

	return msg;
}

/**
 * Returns the name for the selected range if name exists. 
 * @type String
 * @throws Exception if unable to retrieve the named range
 */
MLA.getSelectedRangeName = function()
{
	var msg = window.external.getSelectedRangeName();

	var errMsg = MLA.errorCheck(msg);

	if(errMsg!=null)
        	throw("Error: Not able to getSelectedRangeName; "+errMsg);
        return msg;
}

/**
 * Returns the name for the selected chart. 
 * @type String
 * @throws Exception if unable to retrieve the chart name
 */
MLA.getSelectedChartName = function()
{
	var msg = window.external.getSelectedChartName();

	var errMsg = MLA.errorCheck(msg);

	if(errMsg!=null)
        	throw("Error: Not able to getSelectedChartName; "+errMsg);
        return msg;
}

/**
 * Returns cells selected in active workbook.  This works for contigous cells.  When disparate cells are selected, the last contigous range of cells selected in the active workbook will be returned.
 * @type MLA.Cell 
 * @throws Exception if unable to retrieve the coordinates
 */
MLA.getSelectedCells = function()
{
     var cellresults = window.external.getSelectedCells();

     var errMsg = MLA.errorCheck(cellresults);

     if(errMsg!=null) 
        throw("Error: Not able to getSelectedCells; "+errMsg);

     var cellstring = "{ \"cells\" : "+cellresults+"}";
     //alert (cellstring);
     var cells = eval('('+cellstring+')');
     //alert(cells.cells.length);
     var cellArray = new Array();
     for(var i =0;i<cells.cells.length;i++){
	     var cell = cells.cells[i];
	     var mlacell = new MLA.Cell();
             mlacell=cell;
	     cellArray[i]=mlacell;
	     //alert("coordinate: "+cell.coordinate+"arraylength"+cellArray.length);
	     //alert("cell Coordinate: "+cellArray[i].coordinate+" rowIdx: "+cellArray[i].rowIdx+" colIdx: "+cellArray[i].colIdx+" value2: "+cellArray[i].value2);
     }

     return cellArray;
}

/**
 * Returns active cell from the active worksheet in active workbook.  For any range of selected cells, one will always be identified as active; the last selected cell for any range.
 * @type MLA.Cell
 * @throws Exception if unable to retrieve the active cell
 */
MLA.getActiveCell = function()
{
	var cellinfo = window.external.getActiveCell();
        var errMsg = MLA.errorCheck(cellinfo);

        if(errMsg!=null) 
        	throw("Error: Not able to getActiveCell; "+errMsg);

	var cellValues = cellinfo.split(":");
	var rowIdx = cellValues[0];
	var colIdx = cellValues[1];
	var newCell = new MLA.Cell(parseInt(rowIdx),parseInt(colIdx));

	if(cellValues[2]==""){
		newCell.value2=null;
	}else{
	        newCell.value2 = cellValues[2];
	}

	if(cellValues[3]==""){
		newCell.formula=null;
	}else{
		newCell.formula = cellValues[3];
	}

	return newCell;
}

/** @ignore */
MLA.getActiveCellRange = function()
{
	var cell = window.external.getActiveCellRange();
        var errMsg = MLA.errorCheck(cell);

        if(errMsg!=null) 
        	throw("Error: Not able to getActiveCellRange; "+errMsg);
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
	var errMsg = MLA.errorCheck(cellstring);

        if(errMsg!=null) 
        	throw("Error: Not able to getActiveCellText; "+errMsg);

	var cell = eval('('+cellstring+')');
	var mlaCell = new MLA.Cell();
	mlaCell = cell;
	return mlaCell;
}

/**
 * Set the value for the cell at the current cursor position.
 * @param value the value to be inserted in the active cell
 * @throws Exception if unable to set the active cell text
 */
MLA.setActiveCellValue = function(value)
{
	var msg = window.external.setActiveCellValue(value);
	var errMsg = MLA.errorCheck(msg);

        if(errMsg!=null) 
        	throw("Error: Not able to setActiveCellValue; "+errMsg);

        return msg;
}

/**
 * Sets the values for the cells identified by Cell.coordinate.
 * @param cells an array of MLA.Cell objects, the values of which will be used for the values for the given cells in the active workbook.
 * @param sheetname (optional) the name of the worksheet where the Cell values should be populated.  If no sheetname is provided, the cells will be populated in the active worksheet.
 * @throws Exception if unable to set the values for the given cells
 */
MLA.setCellValue = function(cells, sheetname)
{ 
	
	if(sheetname==null)
		sheetname="active";

	//alert("IN FUNCTION");

	var v_array = MLA.isArray(cells);

	if(v_array){
		for(var i =0; i<cells.length; i++){
          		var msg = window.external.setCellValueA1(cells[i].coordinate, cells[i].value2, sheetname);
		        var errMsg = MLA.errorCheck(msg);

                        if(errMsg!=null) 
        	        	throw("Error: Not able to setCellValue; "+errMsg);

		}
	}

	return msg;
}
/**
 * Converts an A1 notation coordinate to R1C1 notation.
 * @param coord the A1 coordinate to be converted
 * @throws Exception if unable to convert the coordinate
 */
MLA.convertA1ToR1C1 = function(coord)
{
	var msg=window.external.convertA1ToR1C1(coord);
	var errMsg = MLA.errorCheck(msg);

        if(errMsg!=null) 
        	throw("Error: Not able to convertA1toR1C1; "+errMsg);

	return msg;
}
/**
 * Converts a row index and column index to an A1 notation coordinate.
 * @param rowIdx the row index 
 * @param colIdx the column index
 * @throws Exception if unable to convert to A1 notation
 */
MLA.convertR1C1ToA1 = function(rowIdx, colIdx)
{
	var msg=window.external.convertR1C1ToA1(rowIdx, colIdx);
	var errMsg = MLA.errorCheck(msg);

        if(errMsg!=null) 
        	throw("Error: Not able to convertR1C1toA1; "+errMsg);

	return msg;
}

/**
 * Clears the contents of the active worksheet in the active workbook.
 * @param sheetName default 'active'
 * @throws Exception if unable to clear the contents of the worksheet specified in the active workbook
 */
MLA.clearWorksheet = function(sheetName)
{
	if(sheetName==null || sheetName==""){
		sheetName="active";
	}

	var msg=window.external.clearWorksheet(sheetName);
	var errMsg = MLA.errorCheck(msg);

        if(errMsg!=null) 
        	throw("Error: Not able to clearWorksheet; "+errMsg);

        return msg;
}

/**
 * Returns the path being used for the /temp dir on the client system.
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

/**
 * Saves the active workbook to MarkLogic.
 * @param tmpPath the path for the /tmp dir on the client system. (have to save a local copy) 
 * @param docTitle the title of the document
 * @param url the url on MarkLogic Server where the XQuery to save can be found
 * @param uname the username for MarkLogic Server
 * @param pwd the password for MarkLogic Server
 * @throws Exception if unable to save the document to MarkLogic
 */
MLA.saveActiveWorkbook = function(tmpPath, doctitle, url, uname,pwd)
{
       var msg = window.external.saveActiveWorkbook(tmpPath, doctitle, url,uname,pwd);
      	var errMsg = MLA.errorCheck(msg);

        if(errMsg!=null) 
        	throw("Error: Not able to saveActiveWorkbook; "+errMsg);

       return msg;
}
/**
 * Open a .xlsx from MarkLogic into Excel.
 * @param tmpPath the path for the /tmp dir on the client system. (have to save a local copy) 
 * @param docuri the uri for the document in MarkLogic Server
 * @param url the url on MarkLogic Server where the XQuery to open the document specified by docuri can be found
 * @param uname the username for MarkLogic Server
 * @param pwd the password for MarkLogic Server
 * @throws Exception if unable to open the document into Excel
 */
MLA.openXlsx = function(tmpPath, docuri, url, uname, pwd)
{
        var msg =  window.external.OpenXlsx(tmpPath, docuri, url, uname,pwd);
	var errMsg = MLA.errorCheck(msg);

        if(errMsg!=null) 
        	throw("Error: Not able to openXlsx; "+errMsg);
        return msg;
}

/**
 * Saves chart in Excel Workbook to filesystem as .png.
 * @param chartPath the path for the /tmp dir on the client system. (have to save a local copy) where you want to save the chart image
 * @type void
 * @throws Exception if unable to save .png to local filesystem
 */
MLA.exportChartImagePNG = function(chartPath)
{
	var msg =  window.external.exportChartImagePNG(chartPath);

	var errMsg = MLA.errorCheck(msg);
        if(errMsg!=null) 
        	throw("Error: Not able to exportChartImagePNG; "+errMsg);
	
        return msg;
}

/**
 * Deletes picture (image), specified by name, from specified sheet in Excel Workbook.
 * @param sheetName the name of the sheet containing picture to be deleted
 * @param imageName the name of the picture to be deleted
 * @type void
 * @throws Exception if unable to delete the picture
 */
MLA.deletePicture = function(sheetName, imageName)
{
        var msg = window.external.deletePicture(sheetName, imageName);

	var errMsg = MLA.errorCheck(msg);
        if(errMsg!=null) 
        	throw("Error: Not able to deletePicture; "+errMsg);
	
        return msg;
}

/**
 * Inserts a base64 encoded string into sheet, specified by sheetName, as an image.
 * @param base64String the base64 string to insert
 * @param sheetName (optional) default 'active'
 * @type void
 * @throws Exception if unable to insert the image
 */
MLA.insertBase64ToImage = function(base64String, sheetName)
{
	if(sheetName==null || sheetName==""){
		sheetName="active";
	}

  	var msg =  window.external.insertBase64ToImage(base64String, sheetName);

	var errMsg = MLA.errorCheck(msg);
        if(errMsg!=null) 
        	throw("Error: Not able to insertBase64ToImage; "+errMsg);
	
        return msg;
}

/**
 * Returns image found at chartPath as base64 encoded string
 * @param chartPath the name of the image, including full path on filesystem, to be encoded 
 * @type String
 * @throws Exception if unable to delete the picture
 */
MLA.base64EncodeImage = function(chartPath)
{
	var msg =  window.external.base64EncodeImage(chartPath);

	var errMsg = MLA.errorCheck(msg);
        if(errMsg!=null) 
        	throw("Error: Not able to base64EncodeImage; "+errMsg);
	
        return msg;
}

/**
 * Deletes file found with full path at parameter filePath, from the filesystem.
 * @param filePath the name of the file to be deleted
 * @type void
 * @throws Exception if unable to delete the file
 */
MLA.deleteFile = function(filePath)
{
	var msg =  window.external.deleteFile(filePath);

	var errMsg = MLA.errorCheck(msg);
        if(errMsg!=null) 
        	throw("Error: Not able to deleteFile; "+errMsg);
	
        return msg;

}

//functions to add/remove events here
/**
 * Adds chartObjectMouseDown events for each chart on sheet specified by sheetName.
 * @param sheetName the name of the sheet to add event listeners to
 * @type void
 * @throws Exception if unable to add chartObjectMouseDown events
 */
MLA.addChartObjectMouseDownEvents = function(sheetName)
{
	var msg =  window.external.addChartObjectMouseDownEvents(sheetName);

	var errMsg = MLA.errorCheck(msg);

	
        if(errMsg!=null) 
        	throw("Error: Not able to addChartObjectMouseDownEvents; "+errMsg);
	
        return msg;
}

/**
 * Removes chartObjectMouseDown events for each chart on sheet specified by sheetName.
 * @param sheetName the name of the sheet to remove event listeners from
 * @type void
 * @throws Exception if unable to remove chartObjectMouseDown events
 */
MLA.removeChartObjectMouseDownEvents = function(sheetName)
{
	var msg =  window.external.removeChartObjectMouseDownEvents(sheetName);

	var errMsg = MLA.errorCheck(msg);
        if(errMsg!=null) 
        	throw("Error: Not able to removeChartObjectMouseDownEvents; "+errMsg);
	
        return msg;
}

/**
 * Returns macro code as text from macro component specified by index for active workbook.
 * @param index the index of the macro to decompile and return as string
 * @type String
 * @throws Exception if unable to get macro text
 */
MLA.getMacroText = function(index)
{
	var source = window.external.getMacroText(index);
	var errMsg = MLA.errorCheck(source);
        if(errMsg!=null) 
        	throw("Error: Not able to getMacroText; "+errMsg);
	return source;
}

/**
 * Runs macro, specified by macro name, in the active workbook.
 * @param macroName the name of the macro to run in the active workbook
 * @type void
 * @throws Exception if unable to run macro
 */
MLA.runMacro = function(macroName)
{
	var msg = window.external.runMacro(macroName);
        var errMsg = MLA.errorCheck(msg);
        if(errMsg!=null) 
        	throw("Error: Not able to runMacro; "+errMsg);
	
        return msg;

}

/**
 * Returns macro name for macro component specified by index for active workbook.
 * @param index the index of the macro component
 * @type String
 * @throws Exception if unable to get macro name
 */
MLA.getMacroName = function(index)
{
        var msg = window.external.getMacroName(index);
        var errMsg = MLA.errorCheck(msg);
	//alert(msg);
        if(errMsg!=null) 
        	throw("Error: Not able to getMacroName; "+errMsg);
	
        return msg;
}

/**
 * Returns macro type for macro component specified by index for active workbook.
 * @param index - the index of the macro component
 * @type String
 * @throws Exception if unable to get macro type 
 */
MLA.getMacroType = function(index)
{
	var msg = window.external.getMacroType(index);
        var errMsg = MLA.errorCheck(msg);
	//alert(msg);
        if(errMsg!=null) 
        	throw("Error: Not able to getMacroType; "+errMsg);
	
        return msg;
}

/**
 * Returns macro procedure name for macro component specified by index for active workbook.
 * @param index - the index of the macro component
 * @type String
 * @throws Exception if unable to get macro name
 */
MLA.getMacroProcedureName = function(index)
{
	var msg = window.external.getMacroProcedureName(index);
        var errMsg = MLA.errorCheck(msg);
        if(errMsg!=null) 
        	throw("Error: Not able to getMacroProcedureName; "+errMsg);
	
        return msg;
}

/**
 * Returns macro comments for macro component specified by index for active workbook. comments here are any that are in the first lines prior to any procedure definition for the macro.
 * @param index the index of the macro component
 * @type String
 * @throws Exception if unable to get macro comments
 */
MLA.getMacroComments = function(index)
{
	var msg = window.external.getMacroComments(index);
        var errMsg = MLA.errorCheck(msg);
        if(errMsg!=null) 
        	throw("Error: Not able to getMacroComments; "+errMsg);
	
        return msg;
}

/**
 * Returns macro signature for macro component specified by index for active workbook. 
 * @param index the index of the macro component
 * @type String
 * @throws Exception if unable to get macro signature
 */
MLA.getMacroSignature= function(index)
{
	var msg = window.external.getMacroSignature(index);
        var errMsg = MLA.errorCheck(msg);
        if(errMsg!=null) 
        	throw("Error: Not able to getMacroSignature; "+errMsg);
	
        return msg;
}

/**
 * Returns count of macros in active workbook. This count can then be used as the index in other functions related functions.
 * @type String
 * @throws Exception if unable to get macro comments
 */
MLA.getMacroCount = function()
{
	var msg=window.external.getMacroCount();
	return msg;
}

/**
 * Adds macro to active workbook.
 * @param source the source code (string) for the macro to be added
 * @param componentType the componentType to add the macro as to the active workbook
 * @type void
 * @throws Exception if unable to add macro to active workbook
 */
MLA.addMacro= function(source, componenType)
{
	var msg = window.external.addMacro(source, componenType);
        var errMsg = MLA.errorCheck(msg);
        if(errMsg!=null) 
        	throw("Error: Not able to addMacro; "+errMsg);
	
        return msg;
}

/**
 * Removes macro, specified by macroName, from active workbook.
 * @param macroName the componentType to add the macro as to the active workbook
 * @type void
 * @throws Exception if unable to remove macro
 */
MLA.removeMacro= function(macroName)
{
	var msg = window.external.removeMacro(macroName);
        var errMsg = MLA.errorCheck(msg);
        if(errMsg!=null) 
        	throw("Error: Not able to removeMacro; "+errMsg);
	
        return msg;
}

