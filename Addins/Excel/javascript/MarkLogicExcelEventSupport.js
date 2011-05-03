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



workbookOpen = function(workbookName)
{
        try{

	    var msg = workbookName;
	    //workbookOpenHandler(sheetName);
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
