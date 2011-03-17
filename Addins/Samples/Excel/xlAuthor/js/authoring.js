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
*/


var SERVER = "http://localhost:8023/pptAuthor";
var USER="oslo";
var AUTH="oslo";

$(document).ready(function() {

       //SET DEFAULTS		
       //by default, current doc selected

       $('#metadata').hide();
       $('#search').hide();
       $('#macro').hide();
       $('#properties').hide();
       $('#noproperties').hide();

   
       //by default,  presentation tags selected
       $('#slidetags').hide();
       $('#shapetags').hide();

       //by default don't show search filtes
       $('#searchfilter').hide();

       //BEGIN top icon tab navigation selection	
       //display current doc tab
       $('a#icon-xl').click(function() {

          $('#main').css('overflow', 'hidden');
	  //css
	  $("#icon-metadata").removeClass("fronticon");
	  $("#icon-search").removeClass("fronticon");
	  $("#icon-macro").removeClass("fronticon");
          $("#icon-xl").addClass("fronticon");     
   
	  //action
          $('#metadata').hide();
          $('#search').hide();
          $('#macro').hide();
          $('#current-doc').show();

	  $('#docnames').empty();

	  $('#metadataForm').children('div').remove();

	  refreshPropertiesPanel();
   
          return false;
  
       });

       //display metadata icon tab
       $('a#icon-metadata').click(function() {
         
          $('#main').css('overflow', 'hidden');

	  //css
	  $("#icon-xl").removeClass("fronticon");
	  $("#icon-search").removeClass("fronticon");
	  $("#icon-macro").removeClass("fronticon");
          $("#icon-metadata").addClass("fronticon");     
   
	  //action 
          $('#current-doc').hide();
          $('#search').hide();
          $('#macro').hide();
          $('#metadata').show();
     
	  $('#docnames').empty();
           
	  refreshTagTree();

          return false;
  
       });

       //display search icon tab
       $('a#icon-search').click(function() {

          $('#main').css('overflow', 'auto');
 	  //css
	  $("#icon-xl").removeClass("fronticon");
	  $("#icon-metadata").removeClass("fronticon");
	  $("#icon-macro").removeClass("fronticon");
          $("#icon-search").addClass("fronticon");     
   
	  //action  
          $('#current-doc').hide();
          $('#metadata').hide();
          $('#macro').hide();
          $('#search').show();

	  $('#docnames').empty();

	  $('#metadataForm').children('div').remove();
   
          return false;
  
       });

       $('a#icon-macro').click(function() {

          //$('#main').css('overflow', 'auto');
 	  //css
	  $("#icon-xl").removeClass("fronticon");
	  $("#icon-metadata").removeClass("fronticon");
          $("#icon-search").removeClass("fronticon");     
	  $("#icon-macro").addClass("fronticon");
   
	  //action  
          $('#current-doc').hide();
          $('#metadata').hide();
          $('#search').hide();
          $('#macro').show();

	  $('#docnames').empty();

          clearMacroMetadataForm();
	  refreshMacroList();

	  //$('#metadataForm').children('div').remove();
   
          return false;
  
       });

       //END top icon tab navigation selection
       
       //BEGIN tag type selection
       //display xlsx tag palette 
       $('a#icon-xlctrl').click(function() {

 	  //css
	  $("#icon-sheetctrl").removeClass("selectedctrl");
	  $("#icon-namedrangectrl").removeClass("selectedctrl");
          $("#icon-xlctrl").addClass("selectedctrl");     
   
	  //action
          $('#slidetags').hide();
          $('#shapetags').hide();
          $('#presentationtags').show();
          setWorkbookProperties();
	   
          return false;
  
       });

       //sheet tag palette
       $('a#icon-sheetctrl').click(function() {   

 	  //css
	  $("#icon-xlctrl").removeClass("selectedctrl");
	  $("#icon-namedrangectrl").removeClass("selectedctrl");
          $("#icon-sheetctrl").addClass("selectedctrl");     
   
	  //action
          $('#presentationtags').hide();
          $('#shapetags').hide();
          $('#slidetags').show();
	  setWorksheetProperties();
   
          return false;
  
       });

       //component tag palette
       $('a#icon-namedrangectrl').click(function() {   

 	  //css
	  $("#icon-xlctrl").removeClass("selectedctrl");
	  $("#icon-sheetctrl").removeClass("selectedctrl");
          $("#icon-namedrangectrl").addClass("selectedctrl");     
   
	  //action
          $('#presentationtags').hide();
          $('#slidetags').hide();
          $('#shapetags').show();
          setComponentProperties();
          return false;
  
       });

       //END tag type selection 
       

       //BEGIN Metadata panel type section
       //display xlsx tag tab
       $('a#icon-meta-xlctrl').click(function() {

 	  //css
	  $("#icon-meta-sheetctrl").removeClass("selectedctrl");
	  $("#icon-meta-namedrangectrl").removeClass("selectedctrl");
          $("#icon-meta-xlctrl").addClass("selectedctrl");     
   
	  //action
	  clearMetadataForm();
	  refreshTagTree();
	   
          return false;
  
       });

       //sheet tag tab
       $('a#icon-meta-sheetctrl').click(function() {

 	  //css
	  $("#icon-meta-xlctrl").removeClass("selectedctrl");
	  $("#icon-meta-namedrangectrl").removeClass("selectedctrl");
          $("#icon-meta-sheetctrl").addClass("selectedctrl");     
   
	  //action
	  clearMetadataForm();
	  refreshTagTree();

          return false;
  
       });

       //component tag tab
       $('a#icon-meta-namedrangectrl').click(function() {

 	  //css
	  $("#icon-meta-xlctrl").removeClass("selectedctrl");
	  $("#icon-meta-sheetctrl").removeClass("selectedctrl");
          $("#icon-meta-namedrangectrl").addClass("selectedctrl");     
   
	  //action
	  clearMetadataForm();
	  refreshTagTree();

          return false;
  
       });
       //END   Metadata panel type section
       
      
       //Blur ctrlbuttons 
       /*  $('#textcontrols').click(function() {
		      
          $("#buttongroup").li.a.blur();
          return false;
        });
       */

       //search form related
       $('#ddbtn').click(function() {
          if( $('#searchfilter').is(':visible')){
             $("#ddbtn").removeClass("ddbtnactive");
             $('#searchfilter').hide();
	  }
	  else{
             $("#ddbtn").addClass("ddbtnactive");
             $('#searchfilter').show();
	  }
       });

       $('#fbtn').click(function() {
          if ($('#fbtn').is('.fbtnactive')){
            $('#fbtn').removeClass("fbtnactive");
	  }
	  else{
            $('#fbtn').addClass("fbtnactive");
	  }
       });

       $('#ML-Message').hide();
       $('#ML-Message2').hide();

//REMOVE COMMENT commenting out for testing in IE, uncomment for release
//refreshPropertiesPanel();
       
});

//for v2:
//listen for onsubmit event and cancel it instead of onkeypress
//avoids issue of numberpad vs. paste, etc.
function checkForEnter()
{
     if (window.event && window.event.keyCode == 13){
	 return searchAction();
     }

     return true;
}


function searchAction(startidx)
{
	if(startidx==null){
	   startidx = 0;
	}


        var cbsx = []; //will contain all checkboxes checked status
        var cbsid = []; //will contain all checkboxes ids

	if ($('#fbtn').is('.fbtnactive')){
          $('#searchfilter input:checkbox').each(function(){
	     if(this.checked){
              cbsid.push(this.id);
              cbsx.push(this.checked);
	     }
          });
	}

	var qry = $('#searchbox').val();
	var searchType =$("#searchtype input[@name='search:bst']:checked").val();
        //alert("query: "+qry+ " startidx: "+startidx+" cbsid: "+cbsid+" searchtype: "+searchType);
	simpleAjaxSearch(qry , startidx, cbsid, searchType);
}

function simpleAjaxSearch(searchval, startidx, cbsid,searchType)
{
    var newurl = "";

    if(startidx==0)
	    newurl = "search/search.xqy";
    else
	    newurl = "search/search.xqy?start="+startidx;

    $.ajax({
          type: "GET",
          url: newurl, //"search/search.xqy",
          data: { qry : searchval, stype : searchType,  params : cbsid },
          success: function(msg){
			try{
                            $('#searchresults').empty();
                            $('#searchresults').append(msg);
                            $('#searchresults').html(msg);
			}catch(e){
			    //v2
 			    //improve error handling error message | display nicely
			    alert("ERROR in simpleAjaxSearch: "+e.description);
			}
                   }
     });
}

function runMacro(macroName)
{
	try{
	   MLA.runMacro(macroName);
	}catch(err){
           alert(err.description);
	}
}

function removeMacro(macroName)
{
	try{
	   MLA.removeMacro(macroName);
	   //remove associated metadata here?  or modify, maintaing some XML for audit?
	   var metadataPartID = getMetadataPartID(macroName);
	   MLA.deleteCustomXMLPart(metadataPartID);
	   alert("macro component removed from workbook");
	}catch(err){
           alert(err.description);
	}
}

function insertMacroAction(contentURL, buttonIndex)
{
	try{
	     simpleAjaxMacroInsert(contentURL, buttonIndex);
	}catch(err){
		alert("ERROR in insertMacroAction(): "+err.description);
	}

}

function insertComponentAction(contenturl, rId, other, buttonIndex)
{
	try{
             if(rId == null) 
		  rId = "";
             //have to pass buttonIndex as insertedPart may not be inserted
	     //when we go to construct the undo button
	     simpleAjaxComponentInsert(contenturl,rId, other, buttonIndex);
	}catch(err){
	     alert("ERROR in insertComponentAction(): "+err.description);
	}
}


function simpleAjaxMacroInsert(contentURL, buttonIndex)
{ 
    $.ajax({
          type: "GET",
          url: "search/insert-macro.xqy",
          data: "uri=" + contentURL,
          success: function(msg){
			try{
			 insertMacroContent(msg, buttonIndex);
			}catch(e){
			  alert("ERROR in SimpleAjaxMacroInsert(): "+e.description);
			}
                   }
    });
}


function simpleAjaxComponentInsert(contenturl,rId, other, buttonIndex)
{ 
    $.ajax({
          type: "GET",
          url: "search/insert-component.xqy",
          data: "uri=" + contenturl+"&rid="+ rId,
          success: function(msg){
			try{
			 insertComponentContent(msg, other, buttonIndex);
			}catch(e){
			  alert("ERROR in SimpleAjaxComponentInsert(): "+e.description);
			}
                   }
    });
}


function insertImage(picuri)
{
       var config = MLA.getConfiguration();
       var fullurl= config.url;
       var picuri = fullurl + "/search/download-support.xqy?uid="+picuri;
//alert("IN FUNCTION" +config + " / " + fullurl +" / "+picuri);
       var msg = MLA.insertImage(picuri,USER,AUTH);
}

function setPictureFormat(pictureFormat)
{
	var slideindex = MLA.getSlideIndex();
	var shapename = MLA.getShapeRangeName();
	var jsonPicFormat = MLA.jsonStringify(pictureFormat);
	var msg = MLA.setPictureFormat(slideindex, shapename, jsonPicFormat);
}

function setUndoButton(buttonIndex,newShapeName)
{
	try
	{
	     var searchType =$("#searchtype input[@name='search:bst']:checked").val();
	     var slideIndex = MLA.getSlideIndex();
	     var shapeName = newShapeName; //MLA.getShapeRangeName();
	     var id = "undobutton"+buttonIndex;
	     var btn = $('#'+id);
	     btn.children('a').remove();
	     btn.append("<a href=\"javascript:undoInsert('"+searchType+"','"+slideIndex+"','"+shapeName+"')\" onmouseup='blurSelected(this)' class='smallbtn'>Undo</a>");
	}catch(err){
               alert("ERROR in setUndoButton(): "+err.description); 
	}

}

function insertMacroContent(content, buttonIndex)
{
 	try{
		var local = MLA.createXMLDOM(content);
		var metaparts = local.getElementsByTagName("dc:metadata"); //one part for now
		var mplength = metaparts.length;
                var macroText = null;
		var macroType = "";
		var macroName = "";
		
		if(mplength > 1){
			alert("length"+mplength);
			//jsonPkg = metaXml.getElementsByTagName("dc:description")[0];
		}
		else{
			macroText = metaparts[0].getElementsByTagName("dc:description")[0].childNodes[0].nodeValue;
			macroType = metaparts[0].getElementsByTagName("dc:type")[0].childNodes[0].nodeValue;
			macroName = metaparts[0].getElementsByTagName("dc:identifier")[0].childNodes[0].nodeValue;
		}


		
		if(macroText==null || macroText ==""){
			//do nothing
		}
		else{
			var mtext = getMacroTextByProcedureName(macroName);
			if(mtext==null || mtext==""){
                          MLA.addMacro(macroText, macroType);
                          for (var i = 0; i < metaparts.length; i++) { 
                            MLA.addCustomXMLPart(metaparts[i].xml);
		          } 

                          alert("Macro component added to workbook.");
			}else{
			  alert("Unable to add macro component to workbook.\n"+
		                "A macro component with that name already exists.");
			}
			//Remove Button set by search.xqy removeMacro(procName);	
                }
	}catch(e){
		alert("ERROR : "+e.description);
	}
}

function insertComponentContent(content, buttonIndex)
{       //take last JSON string for creation of Shape, for each tag on shape, loop through and add custom parts
	//other is a picture currently, but could include more, hence 'other'
	try{
		var local = MLA.createXMLDOM(content);
		var metaparts = local.getElementsByTagName("dc:metadata");
		var mplength = metaparts.length;

		var jsonPkg = null;
		var jsonString = "";
                var shapeRange = null;
		
		if(mplength > 1){
			var metaXml = metaparts[metaparts.length-1];
			jsonPkg = metaXml.getElementsByTagName("dc:description")[0];
		}
		else{
			jsonPkg = metaparts[0].getElementsByTagName("dc:description")[0];
		}

                jsonString = jsonPkg.childNodes[0].nodeValue;
		
		if(jsonString==null || jsonString =="")
		{
			//do nothing
		}
		else
		{
			//alert("HERE");
			shapeRange = MLA.jsonParse(jsonString);
			var newShapeName = "";

			//alert("HERE 2");
			if(shapeRange.shape.type=="msoPicture")
			{
				//alert("HERE 3"+other);
				  insertImage(other);
				//alert("HERE 4"+other);
				  setPictureFormat(shapeRange.pictureFormat);
				//alert("HERE 5"+other);
				  var slideIndex=MLA.getSlideIndex();
				  var newShapeName= MLA.getShapeRangeName();
				  var jsonTags = MLA.jsonStringify(shapeRange.tags);
				//alert("HERE 6"+other);
				  MLA.addShapeTags(slideIndex,newShapeName,jsonTags);
			}
			else
			{
			  var slideIndex = MLA.getSlideIndex();
			      newShapeName = MLA.addShape(slideIndex,shapeRange);
			}

			//tags retain name and value, so now loop through 
			//and just add Custom parts, they're already linked thru tag.value/dc:identifier

                        for (var i = 0; i < metaparts.length; i++) 
			{ 
                          MLA.addCustomXMLPart(metaparts[i].xml);
		        } 

		        setUndoButton(buttonIndex,newShapeName);	
					
                }
	}catch(e)
	{
		alert("error: "+e.description);
	}
}

function blurSelected(btn_element)
{
	btn_element.blur();
}


function getMetadataPartID(ctrlId)
{
	//alert("ctrlId"+ctrlId);
	var customPartIds = MLA.getCustomXMLPartIds();
        var customPartId = null;
	var metadataPartId = null;

        if(customPartIds.length > 0 ){
	   for (i = 0; i < customPartIds.length; i++){
               customPartId = customPartIds[i];
	       var customPart = MLA.getCustomXMLPart(customPartId);
               var id = customPart.getElementsByTagName("dc:identifier")[0];
	       //18649719 - from test doc
	       if(id.childNodes[0].nodeValue==ctrlId){
		  metadataPartId = customPartId;
	       }

	    }
	}

	return metadataPartId;
}

function replaceCustomMetadataPart(partId, metadataPart)
{
	MLA.deleteCustomXMLPart(partId);
	MLA.addCustomXMLPart(metadataPart);
}

function setMetadataPartValues(tagId)
{
	//get id of currently selected control
        var controlID = tagId; //mlacontrolref.id;	

	//get Part ID of Custom XML Part associated with Control
	var metadataPartID = getMetadataPartID(controlID);

	//get Custom XML Part associated with Control
	var metadataPart = MLA.getCustomXMLPart(metadataPartID); 
	var meta = metadataPart.getElementsByTagName("dc:metadata")[0];
        
	//set form values in Custom XML Part
	//start at 2 to skip identifier and first description (we use for json)
	for(var i = 4;i < meta.childNodes.length; i++){
	        var formID="form-"+i+"-"+controlID;
                var value = $('#'+formID).val();
		meta.childNodes[i].text = value;
	}

        //save edited part
	replaceCustomMetadataPart(metadataPartID, meta);

        if( $('#macro').is(':visible')){  //ONLY DO WHEN TREE EXPOSED, MOVE TO EVENT
     	   $('#ML-Message2').show().fadeOut(1500);
        }else if( $('#metadata').is(':visible')) {
	   $('#ML-Message').show().fadeOut(1500);
        }
}

function isScrolledIntoView(ctrlId)
{
    var docViewTop =  $("#treeWindow").scrollTop();
    var docViewBottom = docViewTop +  $("#treeWindow").height();

    var elemTop = $('#'+ctrlId).offset().top + docViewTop;
    var elemBottom = elemTop + $('#'+ctrlId).height();

    //alert("tree top: "+docViewTop+"\ntreebottom: "+docViewBottom+"\n controlltop: "+elemTop+"\n+controlbottom: "+elemBottom);

    var vis = ((elemBottom >= docViewTop) && (elemTop <= docViewBottom)
      && (elemBottom <= docViewBottom) &&  (elemTop >= docViewTop) );

    return vis;
}

function clearMetadataForm()
{
	var form = $('#metadataForm');
        if(form.children('div').length){
		form.children('div').remove();
	}
}

function clearMacroMetadataForm()
{
	//alert("Clearing MacroForm");
	var form = $('#macroMetadataForm');
        if(form.children('div').length){
		form.children('div').remove();
	}
}

function getMacroTypeByProcedureName(procedureName)
{
 	var count = MLA.getMacroCount();
	var type="";

	for(var j=1;j<=count;j++){
		var name = MLA.getMacroProcedureName(j);
		if(name==procedureName){
		   type = MLA.getMacroType(j);
		   break;
		}

  	}

	return type;
}

function getMacroTextByProcedureName(procedureName)
{
 	var count = MLA.getMacroCount();
	var text=null;

	for(var j=1;j<=count;j++){
		var name = MLA.getMacroProcedureName(j);
		if(name==procedureName){
		   text = MLA.getMacroText(j);
		   break;
		}
  	}

	return text;
}

function setMacroFocus(enteredId)
{
	
    if( $('#macro').is(':visible'))  //ONLY DO WHEN TREE EXPOSED, MOVE TO EVENT 
    {
	var tagID = null;
	if(enteredId == null || enteredId == ""){
	        //window.event.cancelBubble=true;
		tagID = window.event.srcElement.id;
	}else{
		tagID = enteredId;
	}
	
         //set highlight of selected using class
	$('#macrolist').find('a').removeClass("selectedtreectrl");
	$('#'+tagID).addClass("selectedtreectrl");

	//clear metaform in panel
	clearMacroMetadataForm();

	var metaform = $('#macroMetadataForm');
	//need to grab custom piece for metadata section
	var macroMetadataID = getMetadataPartID(tagID);

	if(!(macroMetadataID == null)){	
	   var metadata = MLA.getCustomXMLPart(macroMetadataID);
	   var meta = metadata.getElementsByTagName("dc:metadata")[0];

           //start at 4 to skip relation, type, identifier, description
	   for(var i = 4;i < meta.childNodes.length; i++){
		//assumes XML has QName prefix
	        var localname = meta.childNodes[i].nodeName.split(":");
		var formID = "form-"+i+"-"+tagID;
                var child = meta.childNodes[i];	        	
	        var input="";
		var formValue="";

		if(child.childNodes[0] == null){
		        formValue = "";
		}else{
			formValue = child.childNodes[0].nodeValue;
		}

		if(localname[1]=="description"){
			input = "<textarea cols='40' rows='5' wrap='virtual' id='"+formID+"'>"+
				 formValue +
				"</textarea>";
		}
		else{
			input = "<input id='"+formID+"' type='text' value='"+formValue+"'/>";
		}
			
		  metaform.append("<div>"+
		  		     "<p><label>"+localname[1]+"</p></label>"+
				        input+
                                     "<p>&nbsp; </p>"+
				  "</div>");

		  $('#'+formID).change(function() {
                     setMetadataPartValues(tagID); 
                   }); 
	   }//end of for
	}else{
		//addcustomxmlpart, set header values to macro, macrotype, procedurname, macrotext, display rest of of form.  call setMacroFocus(procedurename);
	        
	        var procedureName = tagID;
		var macroType = getMacroTypeByProcedureName(tagID); 
	        var description =  getMacroTextByProcedureName(tagID);
	        //have to get type, macrotext for this macro. 	
		var stringxml = MLA.unescapeXMLCharEntities(generateTemplate(1));
		//var stringxml = MLA.unescapeXMLCharEntities(generateTemplate(map.get('",$value,"')));
                var domxml = MLA.createXMLDOM(stringxml);
                var relation = domxml.getElementsByTagName('dc:relation')[0];
                var type = domxml.getElementsByTagName('dc:type')[0];
                var id = domxml.getElementsByTagName('dc:identifier')[0];
                var desc = domxml.getElementsByTagName('dc:description')[0];

                if(relation.hasChildNodes()){
		     relation.childNodes[0].nodeValue='';
	 	     relation.childNodes[0].nodeValue="macro";
	        }
	        else{
	             var child = relation.appendChild(domxml.createTextNode("macro"));
	        }


	        if(type.hasChildNodes()){
		     type.childNodes[0].nodeValue='';
	 	     type.childNodes[0].nodeValue=macroType;
	        } 
	        else{
	             var child = type.appendChild(domxml.createTextNode(macroType));
	        }

                if(id.hasChildNodes()){
		     id.childNodes[0].nodeValue='';
	 	     id.childNodes[0].nodeValue=procedureName;
	        } 
	        else{
	             var child = id.appendChild(domxml.createTextNode(procedureName));
	        }

                if(desc.hasChildNodes()){
		     desc.childNodes[0].nodeValue='';
	 	     desc.childNodes[0].nodeValue=description;
	        } 
	        else{
	             var child = desc.appendChild(domxml.createTextNode(description));
	        }

                 MLA.addCustomXMLPart(domxml.xml);
		 setMacroFocus(procedureName);

	}	

    }//end if metadata.visible

}

function setTagFocus(enteredId)
{

    if( $('#metadata').is(':visible')){  //ONLY DO WHEN TREE EXPOSED, MOVE TO EVENT 
	var tagID = null;
	if(enteredId == null || enteredId == ""){
	        //window.event.cancelBubble=true;
		tagID = window.event.srcElement.id;
		//alert("tagID: in if "+tagID);
	}
         //set highlight of selected using class
	$('#treelist').find('a').removeClass("selectedtreectrl");
	$('#'+tagID).addClass("selectedtreectrl");


	//clear metaform in panel
	clearMetadataForm();
	var metaform = $('#metadataForm');

	//need to grab custom piece for metadata section
	var metadataID = getMetadataPartID(tagID);

	if(!(metadataID == null)){		
	   var metadata = MLA.getCustomXMLPart(metadataID);
	   var meta = metadata.getElementsByTagName("dc:metadata")[0];

           //start at 2 to skip identifier and first description
	   for(var i = 4;i < meta.childNodes.length; i++){
		//assumes XML has QName prefix
	        var localname = meta.childNodes[i].nodeName.split(":");
		var formID = "form-"+i+"-"+tagID;
                var child = meta.childNodes[i];	        	
	        var input="";
		var formValue="";

		if(child.childNodes[0] == null){
		        formValue = "";
		}else{
			formValue = child.childNodes[0].nodeValue;
		}

		if(localname[1]=="description"){
			input = "<textarea cols='40' rows='5' wrap='virtual' id='"+formID+"'>"+
				 formValue +
				"</textarea>";
		}
		else{
			input = "<input id='"+formID+"' type='text' value='"+formValue+"'/>";
		}
			
		  metaform.append("<div>"+
		  		     "<p><label>"+localname[1]+"</p></label>"+
				        input+
                                     "<p>&nbsp; </p>"+
				  "</div>");

		  $('#'+formID).change(function() {
                     setMetadataPartValues(tagID); 
                   }); 
	   }//end of for
	} 

    }//end if metadata.visible
}

function getWorkbookTags(taggedComponent)
{
	var customPartIds = MLA.getCustomXMLPartIds();
        var customPartId = null;
	var tagArray=[];
        if(customPartIds.length > 0 ){
	   for (i = 0; i < customPartIds.length; i++){
               customPartId = customPartIds[i];
	       var customPart = MLA.getCustomXMLPart(customPartId);
	       //need to add check here for relation
               var  relation = customPart.getElementsByTagName("dc:relation")[0].childNodes[0].nodeValue;
	       if(relation==taggedComponent){
		  var tag =customPart.getElementsByTagName("dc:identifier")[0].childNodes[0].nodeValue;
		  tagArray[i]=tag;
	       }
	   }
	}

	return tagArray;


}

function getWorksheetTags(taggedComponent, taggedType)
{
	var customPartIds = MLA.getCustomXMLPartIds();
        var customPartId = null;
	var tagArray=[];
        if(customPartIds.length > 0 ){
	   for (i = 0; i < customPartIds.length; i++){
               customPartId = customPartIds[i];
	       var customPart = MLA.getCustomXMLPart(customPartId);
	       //need to add check here for relation
               var  relation = customPart.getElementsByTagName("dc:relation")[0].childNodes[0].nodeValue;
               var  type = customPart.getElementsByTagName("dc:type")[0].childNodes[0].nodeValue;
	       if(relation==taggedComponent && type==taggedType){
		  var tag =customPart.getElementsByTagName("dc:identifier")[0].childNodes[0].nodeValue;
		  tagArray[i]=tag;
	       }
	   }
	}

	return tagArray;

}

function getComponentTags(taggedComponent, taggedType, sheetName) //relation, type(name), sheetName
{
//alert("Tagged Component: "+taggedComponent+" Tagged Type "+taggedType+" Sheet Name "+sheetName);

	var customPartIds = MLA.getCustomXMLPartIds();
        var customPartId = null;
	var tagArray=[];
        if(customPartIds.length > 0 ){
	   for (i = 0; i < customPartIds.length; i++){
               customPartId = customPartIds[i];
	       var customPart = MLA.getCustomXMLPart(customPartId);
	       //need to add check here for relation
               var relation = customPart.getElementsByTagName("dc:relation")[0].childNodes[0].nodeValue;
               var type = customPart.getElementsByTagName("dc:type")[0].childNodes[0].nodeValue;
               var tag = customPart.getElementsByTagName("dc:identifier")[0].childNodes[0].nodeValue;
              
	       if(relation==taggedComponent && type==taggedType){
			       tagArray[i]=tag;
	       }
		       
	   }
	}

	return tagArray;


}

function refreshMacroList()
{
     $('#properties').children('div').remove();

     if($('#macrolist').children('li').length){   
	 $('#macrolist').children('li').remove();
	 $('#macrolist').children('ul').remove();
     }

     var myList = $('#macrolist');

     var mcount = MLA.getMacroCount();
     var names =[];
     var types =[];
     if(mcount > 0){
	 for(var j=1;j<=mcount;j++){ 
           var name = MLA.getMacroProcedureName(j); 
	   var type = MLA.getMacroType(j);
            if(type=="vbext_ct_StdModule" || type=="vbext_ct_ClassModule"){
		 names[j]=name;
		 types[j]=type;
	    }
	 }

	 if(names.length>0){       
		 for(var m=0;m<names.length;m++){
		   var mname = names[m];
		   var mtype = types[m];

	           if(mname=="" || mname==null){
		    //do nothing
	           }else{
		   myList.append("<li>"+
		  		   "<a href='#' id='"+mname+"'>"+
				    "<span id='"+mname+"'>"+
				       mname +
                                    "</span>"+
				   "</a>"+
			         "</li>");

		    var dref = $('#'+mname);

	            dref.bind('click', function() {
                        setMacroFocus(); 
                     });
		   }
		 }

	 }
	 else{
		 myList.append("<li>"+
		            "<a href='#' id='nomacros'>"+
				"<span id='nomacros'>"+
				  "The Workbook does not contain additional macros."+
                                "</span>"+
			     "</a>"+
			    "</li>");
	 }

     }
     


}

function refreshTagTree()
{
     //didn't have to do this with word as we cleared them with an event

     $('#properties').children('div').remove();

     if($('#treelist').children('li').length){   
	 $('#treelist').children('li').remove();
	 $('#treelist').children('ul').remove();
     }

     var myList = $('#treelist');

     //var docTitle = MLA.getActiveWorkbookName();

     if( $('#icon-meta-xlctrl').is('.selectedctrl')){

        var wb_tags = getWorkbookTags("workbook");
 
	//var iconType = "textIcon";

        if(wb_tags.length==0){
			    myList.append("<li>"+
		            "<a href='#' id='nowbtags'>"+
				"<span id='nowbtags'>"+
				  "The Workbook has not been tagged."+
                                "</span>"+
			     "</a>"+
			    "</li>");
	}else{
	  for(var i =0;i<wb_tags.length;i++){
	    var value = wb_tags[i];

	    if(value=="" || value==null){
		    //do nothing
	    }else{
	
	    myList.append("<li>"+
		            //"<a href=\"javascript:setNewTagFocus('"+value+"')\" id='"+value+"'>"+
		            "<a href='#' id='"+value+"'>"+
				"<span id='"+value+"'>"+
				     value +
                                "</span>"+
			     "</a>"+
			    "</li>");

	    }

	    var aref = $('#'+value);

	    aref.bind('click', function() {
                        setTagFocus(); 
                     });
	  }
	}
	

     }else if( $('#icon-meta-sheetctrl').is('.selectedctrl')){
	var sheetName = MLA.getActiveWorksheetName();
        var ws_tags = getWorksheetTags("worksheet", sheetName);
 
	//var iconType = "textIcon";

        if(ws_tags.length==0){
			    myList.append("<li>"+
		            "<a href='#' id='nowstags'>"+
				"<span id='nowstags'>"+
				  "The Worksheet has not been tagged."+
                                "</span>"+
			     "</a>"+
			    "</li>");
	}else{
	  for(var i =0;i<ws_tags.length;i++){
	    var value = ws_tags[i];

	    if(value=="" || value==null){
		    //alert("NULL"+value);
	    }else{
	
	    myList.append("<li>"+
		            //"<a href=\"javascript:setNewTagFocus('"+value+"')\" id='"+value+"'>"+
		            "<a href='#' id='"+value+"'>"+
				"<span id='"+value+"'>"+
				     value +
                                "</span>"+
			     "</a>"+
			    "</li>");


	    var bref = $('#'+value);

	    bref.bind('click', function() {
                        setTagFocus(); 
                     });
	    }
	  }
	}
	

     }else{  //a namedrange or chart
        var componentName = "";
        var componentRelation = "";
        var chartName = MLA.getSelectedChartName();
        if(!(chartName==null || chartName=="")){
	          componentName=chartName;
	          componentRelation="chart";
        }else{
                  componentName = MLA.getSelectedRangeName();
                  componentRelation = "namedrange";
	}

	if(componentName==null || componentName ==""){
           var tagged = checkForComponentTags()

		if(tagged){
		           myList.append("<li>"+
		            "<a href='#' id='nosheettags'>"+
				"<span id='nocomponentselected'>"+
				  "Worksheet Includes Tagged Components."+
                                "</span>"+
			     "</a>"+
			    "</li>");
		 }else{
			   myList.append("<li>"+
		            "<a href='#' id='nosheettags'>"+
				"<span id='nocomponentselected'>"+
				  "No Components have been Tagged in this Worksheet."+
                                "</span>"+
			     "</a>"+
			    "</li>");
		 }

		 clearMetadataForm();

	     }else{  //check to see if selected component has metadata
		     var sheetName = MLA.getActiveWorksheetName();
                     var comp_tags = getComponentTags(componentRelation, componentName, sheetName);
		     
                     if(comp_tags.length==0){
			    myList.append("<li>"+
		            "<a href='#' id='nosheettags'>"+
				"<span id='nocomponenttags'>"+
				  "The Component has not been tagged."+
                                "</span>"+
			     "</a>"+
			    "</li>");
	             }else{
		     
	               for(var i =0;i<comp_tags.length;i++){
	  	           var value = comp_tags[i];
	                   if(value=="" || value==null){
		                 //alert("NULL"+value);
	                   }else{
	
	                     myList.append("<li>"+
		                             "<a href='#' id='"+value+"'>"+
				                "<span id='"+value+"'>"+
				                     value +
                                                "</span>"+
			                     "</a>"+
			                   "</li>");

	                   }

	                  var cref = $('#'+value);

	                  cref.bind('click', function() {
                                     setTagFocus(); 
                                  });
	               }
		     }
	     }
     }
		    
}

function randomId()
{
    var currentTime = new Date();	
    var randomNum = Math.floor(Math.random()*50000);
    var id =   // currentTime.getHours()+":" +
   	       // currentTime.getMinutes() + ":" +
	       // currentTime.getSeconds() + ":" +
	       "XL"+currentTime.getTime()+randomNum;
 
    return id;
}

function checkWorkbookTags(value)
{
    //var pres_tags = MLA.getPresentationTags();

    var same = false;

    	var customPartIds = MLA.getCustomXMLPartIds();
        var customPartId = null;

        if(customPartIds.length > 0 ){
	   for (i = 0; i < customPartIds.length; i++){
               customPartId = customPartIds[i];
	       var customPart = MLA.getCustomXMLPart(customPartId);
	       //need to add check here for relation
               var relation = customPart.getElementsByTagName("dc:relation")[0].childNodes[0].nodeValue;
               var id = customPart.getElementsByTagName("dc:identifier")[0].childNodes[0].nodeValue;
	       if(relation=="workbook"){
	         if(id==value){
		  //alert("This tag has been applied");
		  same = true;
	         }
	       }
	    }
	}
	
    return same;
}


function checkWorksheetTags(value)
{
	var same = false;

    	var customPartIds = MLA.getCustomXMLPartIds();
        var customPartId = null;

        if(customPartIds.length > 0 ){
	   for (i = 0; i < customPartIds.length; i++){
               customPartId = customPartIds[i];
	       var customPart = MLA.getCustomXMLPart(customPartId);
               var relation = customPart.getElementsByTagName("dc:relation")[0].childNodes[0].nodeValue;
               var id = customPart.getElementsByTagName("dc:identifier")[0].childNodes[0].nodeValue;
	       if(relation=="worksheet"){
	         if(id==value){
       			 same = true;
	         }
	       }
	    }
	}
	
    return same;

}

function checkComponentTags(value)
{
	//can't have 2 named ranges with same name in a sheet
	//so, we'll also assume you can't have 2 charts with same name in a sheet
	//type (used as name of component) will be the unique identifier
	//
	//go back, and if namedrange exists, do a delete/add?
	//alert("In Check Component Tags"+value);
	var same = false;
        var componentName = "";
        var componentRelation = "";
	var coords ="";
	var description ="";
	
                         
              var chartName = MLA.getSelectedChartName();
              if(!(chartName==null || chartName=="")){
		   componentName=chartName;
		   componentRelation="chart";
		   //alert("A CHART"+componentName);
                   //var sheetName = MLA.getActiveWorksheetName();
              }else{
                   componentName = MLA.getSelectedRangeName();
		   coords = MLA.getSelectedRangeCoordinates();
                   componentRelation = "namedrange";
		   //alert("A NAMED RANGE"+componentName);
	      }

	      if(!(componentName==null || componentName =="")){
		     // alert("componentName in IF: "+componentName);
             	  var customPartIds = MLA.getCustomXMLPartIds();
                  var customPartId = null;

                  if(customPartIds.length > 0 ){
	            for (i = 0; i < customPartIds.length; i++){
                        customPartId = customPartIds[i];
	                var customPart = MLA.getCustomXMLPart(customPartId);
                        var relation = customPart.getElementsByTagName("dc:relation")[0].childNodes[0].nodeValue;
                        var type = customPart.getElementsByTagName("dc:type")[0].childNodes[0].nodeValue;
                        var id = customPart.getElementsByTagName("dc:identifier")[0].childNodes[0].nodeValue;

	                 if(relation=="namedrange"){ 
                            description = customPart.getElementsByTagName("dc:description")[0].childNodes[0].nodeValue;
		            if(description==coords){
				    alert("A named range for the selected coordinates already exists.");
				    same=true;
			    }else if(type==componentName){
		            //alert("componentName "+componentName);
	                           if(id==value){
		                     alert("The name entered already exists. Enter a unique name.");
		                     same = true;
	                           }
			    }
	                 }else if(relation=="chart"){
			 var sheetName = MLA.getActiveWorksheetName();
			     if(startsWith(type,sheetName) && id==value){
				alert("The name entered already exists. Enter a unique name.");
				same = true;
			     }
			 }
	            }
	          }

	      }else{
		      //now check no namedrange type in customxml parts   ends with value or starts with the active sheetname
		      //applying same label to different named range won't work.  they have to be unique on a sheet, this is MS
		      //the namedrange will move, and you'll end up with 1 nr in the sheet, but 2 customxml parts with the same relation/type/identifier
                   //alert("IN THE ELSE"); 
                  var sheetName = MLA.getActiveWorksheetName();
		  var customPartIds = MLA.getCustomXMLPartIds();
                  var customPartId = null;
                  if(customPartIds.length > 0 ){
	            for (i = 0; i < customPartIds.length; i++){
                        customPartId = customPartIds[i];
	                var customPart = MLA.getCustomXMLPart(customPartId);
                        var type = customPart.getElementsByTagName("dc:type")[0].childNodes[0].nodeValue;
                        var description = customPart.getElementsByTagName("dc:description")[0].childNodes[0].nodeValue;
			//alert("type: "+type+ " description: "+description+" coords: "+coords+" sheetName: "+sheetName);
			  if(startsWith(type,sheetName) && endsWith(type,value)){
				alert("The name entered already exists. Enter a unique name.");
				same = true;
			  }
		    }
		  }

	      }

    return same;
}


function deleteCustomPart(relation, type, partId)
{
    var customPieceIds = MLA.getCustomXMLPartIds();
    var customPieceId = null;

    if(customPieceIds.length > 0 ){
	for (i = 0; i < customPieceIds.length; i++){
            customPieceId = customPieceIds[i];
	    var customPiece = MLA.getCustomXMLPart(customPieceId);
            var relationX = customPiece.getElementsByTagName("dc:relation")[0].childNodes[0].nodeValue;
            var typeX = customPiece.getElementsByTagName("dc:type")[0].childNodes[0].nodeValue;
            var idX = customPiece.getElementsByTagName("dc:identifier")[0].childNodes[0].nodeValue;
	    //18649719 - from test doc
	    if(relationX==relation){
	     if(typeX==type){
	      if(idX==partId){
		 //alert("deleting part: "+customPieceId);
	         MLA.deleteCustomXMLPart(customPieceId);
	      }
	     }
	    }
	}
    }
}


function deleteWorkbookTag(relation, type, tagvalue)
{
    deleteCustomPart(relation, type, tagvalue);
    setWorkbookProperties();
}

function deleteWorksheetTag(relation, type, tagvalue)
{
    deleteCustomPart(relation, type, tagvalue);
    setWorksheetProperties();
}

function deleteComponentTag(relation, type, tagvalue)
{
    if(relation=="namedrange")
	    MLA.removeNamedRange(type);

    deleteCustomPart(relation, type, tagvalue);
    setComponentProperties();
}


function setWorkbookProperties()
{

    var presProps = $('#properties');
    presProps.children('div').remove();
    $('#properties').show();
    $('#noproperties').hide();

   // var pres_tags = MLA.getPresentationTags();
        var tagHtml = "";
        var customPartIds = MLA.getCustomXMLPartIds();
        var customPartId = null;

        if(customPartIds.length > 0 ){
	   for (i = 0; i < customPartIds.length; i++){
               customPartId = customPartIds[i];
	       var customPart = MLA.getCustomXMLPart(customPartId);
               var id = customPart.getElementsByTagName("dc:relation")[0];
	       if(id.childNodes[0].nodeValue=="workbook"){
		  var relation =customPart.getElementsByTagName("dc:relation")[0].childNodes[0].nodeValue;
		  var type =customPart.getElementsByTagName("dc:type")[0].childNodes[0].nodeValue;
		  var tag =customPart.getElementsByTagName("dc:identifier")[0].childNodes[0].nodeValue;
		  //alert("This tag has been applied");
		  	tagHtml += "<a href=\"javascript:deleteWorkbookTag('"+relation+"','"+type+"','"+tag+"')\" id='"+tag+"'>"+
		      "<span class='deleteIcon' title='delete tag'><strong><label>"+tag+"</label></strong></span>"+
	           "</a>"+
	           "<br/>";
	       }

	    }
	}

	//delete icon will use tag.value for customxml part,tag.name for tag


    if(tagHtml==""){
	presProps.append("<div id='properties'>Workbook has no Tags</div>");
    }else{
	presProps.append("<div id='properties'>Workbook has been tagged:<br/><br/>"+tagHtml+"</div>");
    }
	
}

function setWorksheetProperties()
{
    var presProps = $('#properties');
    presProps.children('div').remove();
    $('#properties').show();
    $('#noproperties').hide();

        var tagHtml = "";
        var customPartIds = MLA.getCustomXMLPartIds();
        var customPartId = null;
	var sheetName = MLA.getActiveWorksheetName();

        if(customPartIds.length > 0 ){
	   for (i = 0; i < customPartIds.length; i++){
               customPartId = customPartIds[i];
	       var customPart = MLA.getCustomXMLPart(customPartId);
	       var relation =customPart.getElementsByTagName("dc:relation")[0].childNodes[0].nodeValue;
	       var type =customPart.getElementsByTagName("dc:type")[0].childNodes[0].nodeValue;
	       var tag =customPart.getElementsByTagName("dc:identifier")[0].childNodes[0].nodeValue;
	       if(relation=="worksheet"){
		       if(type==sheetName){
		  	tagHtml += "<a href=\"javascript:deleteWorksheetTag('"+relation+"','"+type+"','"+tag+"')\" id='"+tag+"'>"+
		      "<span class='deleteIcon' title='delete tag'><strong><label>"+tag+"</label></strong></span>"+
	           "</a>"+
	           "<br/>";
		       }
	       }

	    }
	}

	//delete icon will use tag.value for customxml part,tag.name for tag


    if(tagHtml==""){
	presProps.append("<div id='properties'>Worksheet has no Tags</div>");
    }else{
	presProps.append("<div id='properties'>Worksheet has been tagged:<br/><br/>"+tagHtml+"</div>");
    }

}

function undoInsert(searchType, slideIndex, shapeName)
{
    //var searchType =$("input[@name='search:bst']:checked").val();
    //var slideIndex = MLA.getSlideIndex();
    if(searchType == null || searchType ==""){
	alert("You must first insert content before you can undo it");
    }else{
	if(searchType == "slide"){
		var slide_tags = MLA.getSlideTags(slideIndex);
                for(var i =0;i<slide_tags.tags.length;i++){
	            var tag = slide_tags.tags[i];
		    var tagName = tag.name;
	            var tagValue = tag.value;
		    MLA.deleteSlideTag(slideIndex, tagName);
	            deleteCustomPart(tagValue);

		}

	        var slideShapeNames = MLA.getSlideShapeNames(slideIndex);

		for(var j=0;j<slideShapeNames.length;j++){
		      
		      var shapeName = slideShapeNames[j];
		      var shapeRange = MLA.getShapeRangeView(slideIndex, shapeName);
                      var shape_tags = shapeRange.tags;
	              for(var i =0;i<shape_tags.length;i++){
		         var tag = shape_tags[i];
	                 var tagName = tag.name;
		         var tagValue = tag.value;
		         MLA.deleteShapeTag(slideIndex, shapeName, tagName);
	                 deleteCustomPart(tagValue);
		      }

		}
	  	
		var msg = MLA.deleteSlide(slideIndex);
	}
	else
	{
		//var shapeName = MLA.getShapeRangeName();
                var shapeRange = MLA.getShapeRangeView(slideIndex, shapeName);
                var shape_tags = shapeRange.tags;
	        for(var i =0;i<shape_tags.length;i++){
		    var tag = shape_tags[i];
	            var tagName = tag.name;
		    var tagValue = tag.value;
		    MLA.deleteShapeTag(slideIndex,shapeName, tagName);
	            deleteCustomPart(tagValue);
		}

		var msg = MLA.deleteShape(slideIndex, shapeName);
	     
	}
    }
}

function checkForComponentTags()
{
	//update for charts
	var tagged = false;
	//have to see if sheet includes any tagged charts or named ranges
	var sheetName = MLA.getActiveWorksheetName();
	var namedRanges = MLA.getWorksheetNamedRangeNames(sheetName);
	var chartNames = MLA.getWorksheetChartNames(sheetName);
	
        var customPartIds = MLA.getCustomXMLPartIds();
        var customPartId = null;
        var nr = "";
	var cn = "";

        if(customPartIds.length > 0 && namedRanges.length > 0 ){
           for(var j =0; j< namedRanges.length; j++){
	      if(tagged==false){
	        nr=namedRanges[j];
	        for (var i = 0; i < customPartIds.length; i++){
		    customPartId = customPartIds[i];
	            var customPart = MLA.getCustomXMLPart(customPartId);
		    var type =customPart.getElementsByTagName("dc:type")[0].childNodes[0].nodeValue;
		    if(type==nr){
			    tagged=true;
			    break;
		    }

	        }
	      }
	   }
	}
	else if(customPartIds.length > 0  &&  chartNames.length > 0){
	   //alert("IN THE ELSEIF");

           for(var j =0; j< chartNames.length; j++){
	      if(tagged==false){
	        cn=chartNames[j];
		//alert("CHART NAME"+cn);
	        for (var i = 0; i < customPartIds.length; i++){
		    customPartId = customPartIds[i];
	            var customPart = MLA.getCustomXMLPart(customPartId);
		    var type =customPart.getElementsByTagName("dc:type")[0].childNodes[0].nodeValue;
		    if(startsWith(type,sheetName) && endsWith(type,cn)){
			    //alert("MATCH");
			    tagged=true;
			    break;
		    }

	        }
	      }
	   }
	}

	return tagged;

}

//relation is either workbook, worksheet, component, macro
//type is name of workbook, worksheet, componenet, macrotype (vbe_ct*)
//identifier is tag applied or procedurename
//description is used for chart serizization and macro serialization

////can't have 2 named ranges with same name in a sheet
//so, we'll also assume you can't have 2 charts with same name in a sheet
//type (used as name of component) will be the unique identifie
function setComponentProperties()
{


    var componentProps = $('#properties');
    componentProps.children('div').remove();
    $('#properties').show();
    $('#noproperties').hide();
	
    var componentName = "";
    var componentRelation = "";
    try{                     
          var chartName = MLA.getSelectedChartName();
          if(!(chartName==null || chartName=="")){
	       componentName=chartName;
	       componentRelation="chart";
          }else{
               componentName = MLA.getSelectedRangeName();
               componentRelation = "namedrange";
	  }

	  var tagHtml ="";
	  if(componentName==null || componentName =="")
	  {
             var tagged = checkForComponentTags();

	     if(tagged){
               componentProps.append("<div id='properties'>Worksheet Includes Tagged Components</div>")
	     }else{
	       componentProps.append("<div id='properties'>No Components have been Tagged in this Worksheet</div>");
	     }
	  }else{
               var customPartIds = MLA.getCustomXMLPartIds();
               var customPartId = null;

               if(customPartIds.length > 0 ){
	          for (i = 0; i < customPartIds.length; i++){
                    customPartId = customPartIds[i];
	            var customPart = MLA.getCustomXMLPart(customPartId);
	            //need to add check here for relation
	            var relation =customPart.getElementsByTagName("dc:relation")[0].childNodes[0].nodeValue;
	            var type =customPart.getElementsByTagName("dc:type")[0].childNodes[0].nodeValue;
	            var tag =customPart.getElementsByTagName("dc:identifier")[0].childNodes[0].nodeValue;
	            if(relation=="namedrange" || relation=="chart"){ 
		            if(type==componentName){ 
	                     tagHtml += "<a href=\"javascript:deleteComponentTag('"+relation+"','"+type+"','"+tag+"')\" id='"+tag+"'>"+
		                        "<span class='deleteIcon' title='delete tag'><strong><label>"+tag+"</label></strong></span>"+
	                                "</a>"+
	                                "<br/>";	  
		            }
		     }
	          }

	          if(tagHtml==""){
	             componentProps.append("<div id='properties'>Component has no Tags</div>");
	          }else{
                     componentProps.append("<div id='properties'>Component has been Tagged:<br/><br/>"+tagHtml+"</div>");
	          }

	       }
	  }
 
    }catch(err)
    {
		//do nothing
    }


    return false;
}

function refreshPropertiesPanel()
{
    if( $('#icon-xlctrl').is('.selectedctrl')){
	setWorkbookProperties();
    }else if($('#icon-sheetctrl').is('.selectedctrl')){
	setWorksheetProperties();
    }
    else{
	setComponentProperties();
    }

}

function getJsonString(shapeRange)
{
    var shapeString = "";

    try{
         shapeString = MLA.jsonStringify(shapeRange);
    }
    catch(err){
	 alert(err.description);
    }

    return shapeString;
}

var globalShapeName = "";
var globalSlideIndex = "";

function updateComponentJSON(slideIndex, shapeName)
{

	try{
		   var shapeRangeView = MLA.getShapeRangeView(slideIndex, shapeName);
                  
		   var shape_tags = shapeRangeView.tags;
   
                   for(var i =0;i<shape_tags.length;i++)
                   {
	               var tag = shape_tags[i];
		       var tagId = tag.value;
		       var metadataID = getMetadataPartID(tagId);

                       if(!(metadataID == null))
		       {
	                   var metadata = MLA.getCustomXMLPart(metadataID);

			   var jsonStore = metadata.getElementsByTagName("dc:description")[0];


			   var myShapeString =getJsonString(shapeRangeView);

			   if(jsonStore.hasChildNodes())
			   {
		                 jsonStore.childNodes[0].nodeValue='';
	 	                 jsonStore.childNodes[0].nodeValue=myShapeString;
				 replaceCustomMetadataPart(metadataID, metadata)
			   }

			  // alert(jsonStore.xml);

		       }
		   }
                   alert("Component Information Saved.");

	}catch(err)
	{
		alert("ERROR: "+err.description);
	}
}

function endsWith(str, suffix) {
    return str.indexOf(suffix, str.length - suffix.length) !== -1;
}

function startsWith(str, prefix)
{
    return str.indexOf(prefix) === 0;
}

function trim(str)
{
    return str.replace(/^\s+|\s+$/g, ''); 
}


//BEGIN EVENT HANDLERS
//v2 could define all handlers in application code
//then have application authors tweak, as opposed to editing MarkLogicPowerPointEventSupport.js
function windowSelectionHandler(shapename)
{
    if(!(globalShapeName==shapename))
    {
	var origShapeName = globalShapeName;

	globalShapeName=shapename;

	if($('#icon-shapectrl').is('.selectedctrl'))
	{
	   setComponentProperties();
	}

        if($('#metadata').is(':visible'))
	{  
            clearMetadataForm();
	    if($('#icon-meta-shapectrl').is('.selectedctrl'))
	    {
	       refreshTagTree();
	     
	    }
	}
    }else
    {
            //refreshPropertiesPanel();	
    }
 
    return false;
}

function worksheetSelectionHandler(sheetName)
{
    if( $('#icon-sheetctrl').is('.selectedctrl')){ 
        setWorksheetProperties();
    }
   
    if($('#metadata').is(':visible')){
        refreshTagTree();
	clearMetadataForm();
    }

    if($('#icon-namedrangectrl').is('.selectedctrl')){
	refreshPropertiesPanel();
    }
}
//END EVENT HANDLERS
