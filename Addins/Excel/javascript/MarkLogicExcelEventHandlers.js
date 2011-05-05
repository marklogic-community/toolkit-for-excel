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

MarkLogicExcelEventHandlers.js - javascript api captures events in PowerPoint.

for events you'd like to use, add try/catch and handler call
*/

var debug = false;

MLA.sheetActivateHandler = function(sheetName)
{
        //will need to check, and if in worksheet mode, might highlight something
	//different in pane )(properties, what tags applied to worksheet
	try{
	    var msg = sheetName;
  	    if(debug){
	      alert("in handler sheet activate name: "+sheetName);
	    }

	    var sheetType = MLA.getSheetType(sheetName);
	    //UPDATE UI
	    worksheetSelectionHandler(sheetName);

	    if(debug){
	      alert("sheet type: "+sheetType);
	    }

	    //ADD CHART MOUSEDOWN EVENTS FOR EMBEDDED CHARTS IN SHEET
	    if(sheetType=="xlWorksheet"){
		var msg=MLA.addChartObjectMouseDownEvents(sheetName);
	    }
	}catch(err){
	    msg="error in sheetActivateHandler: "+err.description;
	    alert(msg);
	}

	return msg;
}

MLA.sheetDeactivateHandler = function(sheetName)
{
	try{

	    var msg = sheetName;
            var sheetType = MLA.getSheetType(sheetName);
	    if(sheetType=="xlWorksheet"){
		var msg=MLA.removeChartObjectMouseDownEvents(sheetName);
	    }
	}catch(err){
	    msg="error in sheetDeactivateHandler: "+err.description;
	    alert(msg);
	}
	//alert("in handler sheet deactivate name: "+sheetName);
	return msg;
}

MLA.rangeSelectedHandler = function(rangeName)
{
	try{

	    var msg = rangeName;
	    setComponentProperties();
            if( $('#icon-meta-namedrangectrl').is('.selectedctrl') && $('#icon-meta-namedrangectrl').is(':visible'))
            {
		clearMetadataForm();
	        refreshTagTree();
	    }

	    //need to remove chartObjectEvents here
	}catch(err)
 	{
	    msg="error in rangeSelectedHandler: "+err.description;
	}
	//alert("in handler rangeSelected name: "+rangeName);
	return msg;
}

MLA.chartObjectMouseDownHandler = function(chartName)
{
	var msg=chartName;

	try
	{
           //alert("chartName in handler is "+chartName);
	    if($('#icon-namedrangectrl').is('.selectedctrl') &&  $('#icon-namedrangectrl').is(':visible')){
	        setComponentProperties();
	    }

	    if( $('#icon-meta-namedrangectrl').is('.selectedctrl') && $('#icon-meta-namedrangectrl').is(':visible')){
		clearMetadataForm();
	        refreshTagTree();
	    }

	    //var tmpPath = MLA.getTempPath()+chartName+".PNG"; 
	    //var success = MLA.exportChartImagePNG(tmpPath);
	    //check for error here on success
	    //var base64String=MLA.base64EncodeImage(tmpPath);

	    //if(debug)
	     //alert("CHARTSTRING"+base64String);

	    //var deleted = MLA.deleteFile(tmpPath);

	    //if(debug)
	     //alert("deleted"+deleted);

	}
	catch(err)
	{
	    msg="error in chartObjectMouseDownHandler: "+err.description;
	    alert(msg);

	}
	return msg;
}

MLA.workbookActivateHandler = function(workbookName)
{
   if(debug){
    alert("In workbookActivateHandler, workbookName: "+workbookName);
   }
}

MLA.workbookAfterXmlExportHandler = function(workbookName, mapName, url)
{
   if(debug){
    alert("In workbookAfterXmlExportHandler, workbookName: "+workbookName+" mapName: "+mapName+" url: "+url);
   }
}

MLA.workbookAfterXmlImportHandler = function(workbookName, mapName, refresh)
{
   if(debug){
    alert("In workbookAfterXmlImportHandler, workbookName: "+workbookName+" mapName: "+mapName+" refresh: "+refresh);
   }
}

MLA.workbookBeforeXmlExportHandler = function(workbookName, mapName, url)
{
   if(debug){
    alert("In workbookBeforeXmlExportHandler, workbookName: "+workbookName+" mapName: "+mapName+" url: "+url);
   }
}

MLA.workbookBeforeXmlImportHandler = function(workbookName, mapName, refresh)
{
   if(debug){
    alert("In workbookBeforeXmlImportHandler, workbookName: "+workbookName+" mapName: "+mapName+" refresh: "+refresh);	
   }
}

MLA.workbookBeforeCloseHandler = function(workbookName)
{
   if(debug){
    alert("In workbookBeforeCloseHandler, workbookName: "+workbookName);
   }    
}

MLA.workbookBeforeSaveHandler = function(workbookName)
{
   if(debug){
    alert("In workbookBeforeSaveHandler, workbookName: "+workbookName);	
   }
}

MLA.workbookDeactivateHandler= function(workbookName)
{
   if(debug){
    alert("In workbookDeactivateHandler, workbookName: "+workbookName);	
   }
}

MLA.workbookNewSheetHandler = function(workbookName, sheetName)
{
   if(debug){
    alert("In workbookNewSheetHandler, workbookName: "+workbookName+" sheetName: "+sheetName);
   }
}

MLA.workbookOpenHandler = function(workbookName)
{
   if(debug){
    alert("In workbookOpenHandler, workbookName: "+workbookName);
   }    
}

MLA.sheetBeforeDoubleClickHandler = function(sheetName, range)
{
   if(debug){
    alert("In sheetBeforeDoubleClickHandler, sheetName: "+sheetName+" range"+range);
   }    
}

MLA.sheetBeforeRightClickHandler = function(sheetName, range)
{
   if(debug){
    alert("In sheetBeforeRightClickHandler, sheetName: "+sheetName+" range"+range);
   }    
}

MLA.sheetChangeHandler = function(rangeName)
{
   if(debug){
    alert("In sheetChangeHandler, rangeName: "+rangeName);
   }
}

