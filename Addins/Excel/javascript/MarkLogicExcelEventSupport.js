/* 
Copyright 2008-2011 MarkLogic Corporation

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.

MarkLogicExcelEventSupport.js - javascript api captures events in PowerPoint.

you should NOT edit this file.  to add your own handling, see associated
handlers in MarkLogicExcelEventHandlers.js
*/

var debug = false;

sheetActivate = function(sheetName)
{
	try{
	    var msg = sheetName;
	    MLA.sheetActivateHandler(sheetName);
	}catch(err){
	    msg="error in sheetActivate: "+err.description;
	}

	return msg;
}

sheetDeactivate = function(sheetName)
{
	try{

	    var msg = sheetName;
	    MLA.sheetDeactivateHandler(sheetName);
	}catch(err){
	    msg="error in sheetDeactivate: "+err.description;
	}

	return msg;
}

sheetChange = function(rangeName)
{
  	try{
	    var msg = rangeName;
	    MLA.sheetChangeHandler(rangeName);
	}catch(err){
	    msg="error in sheetChange: "+err.description;
	}

	return msg;
}

//sheetSelectionChange Event, currently only captures event when item selected is a named range
////may update this in the future
rangeSelected = function(rangeName)
{
	try{
	    var msg = rangeName;
	    MLA.rangeSelectedHandler(rangeName);
	}catch(err){
	    msg="error in rangeSelected: "+err.description;
	}

	return msg;
}

workbookActivate = function(workbookName)
{
        try{

	    var msg = workbookName;
	    MLA.workbookActivateHandler(workbookName);
	}catch(err){
            msg="error in workbookActivate: "+err.description;
	}

	return msg;

}
workbookAfterXmlExport = function(workbookName, mapName, url)
{
	try{

	    var msg = workbookName;
	    MLA.workbookAfterXmlExportHandler(workbookName, mapName, url);
	}catch(err){
            msg="error in workbookAfterXmlExport: "+err.description;
	}

	return msg;
}
workbookAfterXmlImport = function(workbookName, mapName, refresh)
{
	try{

	    var msg = workbookName;
	    MLA.workbookAfterXmlImportHandler(workbookName, mapName, refresh);
	}catch(err){
            msg="error in workbookAfterXmlImport: "+err.description;
	}

	return msg;
}

workbookBeforeXmlExport = function(workbookName, mapName, url)
{
	try{
	    var msg = workbookName;
	    MLA.workbookBeforeXmlExportHandler(workbookName, mapName, url);
	}catch(err){
            msg="error in workbookBeforeXmlExport: "+err.description;
	}

	return msg;
}

workbookBeforeXmlImport = function(workbookName, mapName, refresh)
{
	try{
	    var msg = workbookName;
	    MLA.workbookBeforeXmlImportHandler(workbookName, mapName, refresh);
	}catch(err){
            msg="error in workbookBeforeXmlImport: "+err.description;
	}

	return msg;
}

workbookBeforeClose = function(workbookName)
{
        try{
	    var msg = workbookName;
	    MLA.workbookBeforeCloseHandler(workbookName);
	}catch(err){
            msg="error in workbookBeforeClose: "+err.description;
	}

	return msg;

}

workbookBeforeSave = function(workbookName)
{
        try{
	    var msg = workbookName;
	    MLA.workbookBeforeSaveHandler(workbookName);
	}catch(err){
            msg="error in workbookBeforeSave: "+err.description;
	}

	return msg;

}

workbookDeactivate = function(workbookName)
{
        try{
	    var msg = workbookName;
	    MLA.workbookDeactivateHandler(workbookName);
	}catch(err){
            msg="error in workbookDeactivate: "+err.description;
	}

	return msg;

}

workbookNewSheet = function(workbookName, sheetName)
{
        try{
	    var msg = workbookName;
	    MLA.workbookNewSheetHandler(workbookName, sheetName);
	}catch(err){
            msg="error in workbookNewSheet: "+err.description;
	}

	return msg;

}

workbookOpen = function(workbookName)
{
        try{

	    var msg = workbookName;
	    MLA.workbookOpenHandler(workbookName);
	}catch(err){
            msg="error in workbookOpen: "+err.description;
	}

	return msg;

}

chartObjectMouseDown = function(chartName)
{
        try{
	    var msg = chartName;
	    MLA.chartObjectMouseDownHandler(chartName);
	}catch(err){
            msg="error in chartObjectMouseDown: "+err.description;
	}

	return msg;
}
