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
    if(debug){
     alert("in handler sheetActivate sheetName: "+sheetName);
    }
}

MLA.sheetDeactivateHandler = function(sheetName)
{
    if(debug){
     alert("in handler sheetDeactivate sheetName: "+sheetName);
    }
}

MLA.rangeSelectedHandler = function(rangeName)
{
    if(debug){
     alert("in handler rangeSelectedHandler rangeName: "+rangeName);
    }
}

MLA.chartObjectMouseDownHandler = function(chartName)
{
    if(debug){
     alert("in handler chartObjectMouseDownHandler chartName: "+chartName);
    }
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


