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

MarkLogicExcelEventHandlers.js - Excel events captured by MarkLogicExcelEventSupport.js
are routed here. Define your custom handlers here.

*/

var debug = false;

MLA.sheetActivateHandler = function(sheetName)
{
	try{
	    alert("in handler sheetActivate sheetName: "+sheetName);
	}catch(err){
	    msg="error in sheetActivateHandler: "+err.description;
	    alert(msg);
	}

	return msg;
}

MLA.sheetDeactivateHandler = function(sheetName)
{
	try{
	    alert("in handler sheetDeactivate sheetName: "+sheetName);
	}catch(err){
	    msg="error in sheetDeactivateHandler: "+err.description;
	    alert(msg);
	}

	return msg;
}

MLA.rangeSelectedHandler = function(rangeName)
{
	try{
	    alert("in handler rangeSelectedHandler rangeName: "+rangeName);
	}catch(err){
	    msg="error in rangeSelectedHandler: "+err.description;
	    alert(msg);
	}

	return msg;
}
	return msg;
}

MLA.chartObjectMouseDownHandler = function(chartName)
{
	var msg=chartName;

	try
	{
	    alert("in handler chartObjectMouseDownHandler chartName: "+chartName);

	}
	catch(err)
	{
	    msg="error in chartObjectMouseDownHandler: "+err.description;
	    alert(msg);

	}
	return msg;
}


