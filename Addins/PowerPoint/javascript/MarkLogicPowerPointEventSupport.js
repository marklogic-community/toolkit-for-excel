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

MarkLogicPowerPointEventSupport.js - javascript api captures events in PowerPoint.

for events you'd like to use, add try/catch and handler call

version 2.0-1
*/

var debug = false;

windowSelectionChange  = function(shapename)
{
	try{

	var msg = shapename;
	windowSelectionHandler(shapename);
	}
	catch(err)
 	{
		msg="error in windowSelectionChange: "+err.description;
	}

	return msg;
}

windowBeforeRightClick  = function(slideIndex)
{
	try{ 
		var msg = slideIndex;
		if(debug)
	           alert("windowBeforeRightClick: slideIndex "+slideIndex);
	}
	catch(err)
 	{
		msg="error in windowBeforeRightClick: "+err.description;
	}

	return msg;
}

windowBeforeDoubleClick  = function(slideIndex)
{
	try{
		var msg = slideIndex;
	        if(debug)
	          alert("windowBeforeDoubleClick: slideIndex "+slideIndex);
	}
	catch(err)
 	{
		msg="error in windowBeforeDoubleClick: "+err.description;
	}

	return msg;
}

presentationClose  = function(presoName)
{
	try{
		if(debug)
	           alert("presentationClose: presoName "+presoName);
	}
	catch(err)
 	{
		msg="error in presentationClose: "+err.description;
	}

	return msg;
}

presentationSave  = function(presoName)
{
	try{
		if(debug)
	           alert("presentationSave: presoName "+presoName);
	}
	catch(err)
 	{
		msg="error in presentationSave: "+err.description;
	}

	return msg;
}

presentationOpen  = function(presoName)
{
	try{
		if(debug)
	           alert("presentationOpen: presoName "+presoName);
	}
	catch(err)
 	{
		msg="error in presentationOpen: "+err.description;
	}

	return msg;
}

newPresentation  = function(presoName)
{
	try{
		if(debug)
	           alert("newPresentation: presoName "+presoName);
	}
	catch(err)
 	{
		msg="error in newPresentation: "+err.description;
	}

	return msg;
}

presentationNewSlide  = function(slideIndex)
{
	try{
		if(debug)
	           alert("presentationNewSlide: slideIndex "+slideIndex);
	}
	catch(err)
 	{
		msg="error in presentationNewSlide: "+err.description;
	}

	return msg;
}

windowActivate  = function(presoName)
{
	try{
		if(debug)
	           alert("windowActivate: presoName "+presoName);
	}
	catch(err)
 	{
		msg="error in windowActivate: "+err.description;
	}

	return msg;
}

windowDeactivate  = function(presoName)
{
	try{
		if(debug)
	           alert("windowDeactivate: presoName "+presoName);
	}
	catch(err)
 	{
		msg="error in windowDeactivate: "+err.description;
	}

	return msg;
}

slideShowBegin  = function(presoName)
{
	try{
		if(debug)
	           alert("slideShowBegin: presoName "+presoName);
	}
	catch(err)
 	{
		msg="error in slideShowBegin: "+err.description;
	}

	return msg;
}

slideShowNextBuild  = function(presoName)
{
	try{
		if(debug)
	           alert("slideShowNextBuild: presoName "+presoName);
	}
	catch(err)
 	{
		msg="error in slideShowNextBuild: "+err.description;
	}

	return msg;
}

slideShowNextSlide  = function(presoName)
{
	try{
		if(debug)
	           alert("slideShowNextSlide: presoName "+presoName);
	}
	catch(err)
 	{
		msg="error in slideShowNextSlide: "+err.description;
	}

	return msg;
}

slideShowEnd  = function(presoName)
{
	try{
		if(debug)
	           alert("slideShowEnd: presoName "+presoName);
	}
	catch(err)
 	{
		msg="error in slideShowEnd: "+err.description;
	}

	return msg;
}

presentationPrint  = function(presoName)
{
	try{
		if(debug)
	           alert("presentationPrint : slideIndex "+presoName);
	}
	catch(err)
 	{
		msg="error in presentationPrint : "+err.description;
	}

	return msg;
}

slideSelectionChange = function(slideIndex)
{
	try
	{
	   slideSelectionHandler(slideIndex);
	}
	catch(err)
 	{
	    msg="error in slideSelectionChange: "+err.description;
	}
}

colorSchemeChange  = function(shapeRangeName)
{
	try{
		if(debug)
	           alert("colorSchemeChange : shapeRangeName "+shapeRangeName);
	}
	catch(err)
 	{
		msg="error in colorSchemeChange : "+err.description;
	}

	return msg;
}

presentationBeforeSave  = function(presoName)
{
	try{
		if(debug)
	           alert("presentationBeforeSave : presoName "+presoName);
	}
	catch(err)
 	{
		msg="error in presentationBeforeSave : "+err.description;
	}

	return msg;
}

slideShowNextClick  = function(presoName)
{
	try{
		if(debug)
	           alert("slideShowNextClick: presoName "+presoName);
	}
	catch(err)
 	{
		msg="error in slideShowNextClick: "+err.description;
	}

	return msg;
}
