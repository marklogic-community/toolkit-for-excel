/*
Copyright 2008-2010 Mark Logic Corporation

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
       $('#properties').hide();
       $('#noproperties').hide();

   
       //by default,  presentation tags selected
       $('#slidetags').hide();
       $('#shapetags').hide();

       //by default don't show search filtes
       $('#searchfilter').hide();

       //BEGIN top icon tab navigation selection	
       //display current doc tab
       $('a#icon-pptx').click(function() {

          $('#main').css('overflow', 'hidden');
	  //css
	  $("#icon-metadata").removeClass("fronticon");
	  $("#icon-search").removeClass("fronticon");
	  $("#icon-merge").removeClass("fronticon");
          $("#icon-pptx").addClass("fronticon");     
   
	  //action
          $('#metadata').hide();
          $('#search').hide();
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
	  $("#icon-pptx").removeClass("fronticon");
	  $("#icon-search").removeClass("fronticon");
	  $("#icon-merge").removeClass("fronticon");
          $("#icon-metadata").addClass("fronticon");     
   
	  //action 
          $('#current-doc').hide();
          $('#search').hide();
          $('#metadata').show();
     
	  $('#docnames').empty();
           
	  refreshTagTree();

          return false;
  
       });

       //display search icon tab
       $('a#icon-search').click(function() {

          $('#main').css('overflow', 'auto');
 	  //css
	  $("#icon-pptx").removeClass("fronticon");
	  $("#icon-metadata").removeClass("fronticon");
	  $("#icon-merge").removeClass("fronticon");
          $("#icon-search").addClass("fronticon");     
   
	  //action  
          $('#current-doc').hide();
          $('#metadata').hide();
          $('#search').show();

	  $('#docnames').empty();

	  $('#metadataForm').children('div').remove();
   
          return false;
  
       });

       //END top icon tab navigation selection
       
       //BEGIN tag type selection
       //display pptx tag palette
       $('a#icon-pptxctrl').click(function() {

 	  //css
	  $("#icon-slidectrl").removeClass("selectedctrl");
	  $("#icon-shapectrl").removeClass("selectedctrl");
          $("#icon-pptxctrl").addClass("selectedctrl");     
   
	  //action
          $('#slidetags').hide();
          $('#shapetags').hide();
          $('#presentationtags').show();
          setPresentationProperties();
	   
          return false;
  
       });

       //slide tag palette
       $('a#icon-slidectrl').click(function() {

 	  //css
	  $("#icon-pptxctrl").removeClass("selectedctrl");
	  $("#icon-shapectrl").removeClass("selectedctrl");
          $("#icon-slidectrl").addClass("selectedctrl");     
   
	  //action
          $('#presentationtags').hide();
          $('#shapetags').hide();
          $('#slidetags').show();
	  setSlideProperties();
   
          return false;
  
       });

       //shape tag palette
       $('a#icon-shapectrl').click(function() {

 	  //css
	  $("#icon-pptxctrl").removeClass("selectedctrl");
	  $("#icon-slidectrl").removeClass("selectedctrl");
          $("#icon-shapectrl").addClass("selectedctrl");     
   
	  //action
          $('#presentationtags').hide();
          $('#slidetags').hide();
          $('#shapetags').show();
          setShapeProperties();
          return false;
  
       });


       //END tag type selection 
       

       //BEGIN Metadata panel type section
       //display pptx tag palette
       $('a#icon-meta-pptxctrl').click(function() {

 	  //css
	  $("#icon-meta-slidectrl").removeClass("selectedctrl");
	  $("#icon-meta-shapectrl").removeClass("selectedctrl");
          $("#icon-meta-pptxctrl").addClass("selectedctrl");     
   
	  //action
	  clearMetadataForm();
	  refreshTagTree();
	   
          return false;
  
       });

       //slide tag palette
       $('a#icon-meta-slidectrl').click(function() {

 	  //css
	  $("#icon-meta-pptxctrl").removeClass("selectedctrl");
	  $("#icon-meta-shapectrl").removeClass("selectedctrl");
          $("#icon-meta-slidectrl").addClass("selectedctrl");     
   
	  //action
	  clearMetadataForm();
	  refreshTagTree();

          return false;
  
       });

       //shape tag palette
       $('a#icon-meta-shapectrl').click(function() {

 	  //css
	  $("#icon-meta-pptxctrl").removeClass("selectedctrl");
	  $("#icon-meta-slidectrl").removeClass("selectedctrl");
          $("#icon-meta-shapectrl").addClass("selectedctrl");     
   
	  //action
	  clearMetadataForm();
	  refreshTagTree();

          return false;
  
       });
       //END   Metadata panel type section
       

       //END current doc tags selection
      
       //Blur ctrlbuttons 
       /*  $('#textcontrols').click(function() {
		      
          $("#buttongroup").li.a.blur();
          return false;
        });
       */

       //search form related
       $('#ddbtn').click(function() {
          if( $('#searchfilter').is(':visible'))
	  {
             $("#ddbtn").removeClass("ddbtnactive");
             $('#searchfilter').hide();
	  }
	  else
	  {
             $("#ddbtn").addClass("ddbtnactive");
             $('#searchfilter').show();
	  }
       });

       $('#fbtn').click(function() {
          if ($('#fbtn').is('.fbtnactive')) 
            $('#fbtn').removeClass("fbtnactive");
	  else
            $('#fbtn').addClass("fbtnactive");
       });

       $('#ML-Message').hide();

//REMOVE COMMENT commenting out for testing in IE, uncomment for release
//refreshPropertiesPanel();
       
});

//for v2:
//listen for onsubmit event and cancel it instead of onkeypress
//avoids issue of numberpad vs. paste, etc.
function checkForEnter()
{
     if (window.event && window.event.keyCode == 13)
     {
         //alert('Enter key pressed');
	 return searchAction();
     }

     return true;
}


function searchAction(startidx)
{
	if(startidx==null)
	{
	   startidx = 0;
	}


        var cbsx = []; //will contain all checkboxes checked status
        var cbsid = []; //will contain all checkboxes ids

	if ($('#fbtn').is('.fbtnactive'))
	{
        $('#searchfilter input:checkbox').each(function(){
	  if(this.checked){
            cbsid.push(this.id);
            cbsx.push(this.checked);
	    }
        });
	}

	var qry = $('#searchbox').val();
	var searchType =$("#searchtype input[@name='search:bst']:checked").val();

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
			}catch(e)
			{
			    //v2
 			    //improve error handling error message | display nicely
			    alert("ERROR in simpleAjaxSearch: "+e.description);
			}
                   }
     });
}

function insertComponentAction(contenturl, rId, other, buttonIndex)
{
	try{
             if(rId == null) 
		  rId = "";
             //have to pass buttonIndex as insertedPart may not be inserted
	     //when we go to construct the undo button
	     simpleAjaxComponentInsert(contenturl,rId, other, buttonIndex);
	}catch(err)
	{
	     alert("ERROR in insertComponentAction(): "+err.description);
	}
}

function insertSlideAction(contenturl, rId, docuri, slideIdx, buttonIndex)
{
	try{
              var retain="true";
              var tmpPath = MLA.getTempPath();

              var config = MLA.getConfiguration();
              var fullurl= config.url;
              var url = fullurl + "/search/download-support.xqy?uid="+docuri;
	      var slideIndex = slideIdx;  //index of slide to be copied over, not current //MLA.getSlideIndex();
      
              var tokens = docuri.split("/");
              var filename = tokens[tokens.length-1];

              try{
                  var msg = MLA.insertSlide(tmpPath, filename,slideIndex,url,USER,AUTH,retain);
		  //tags are retained, so just check for and insert metadata parts
		  simpleAjaxSlideInsert(contenturl, rId, buttonIndex);

              }catch(e){
                  alert("insertSlideAction Error: "+msg+e);
              }

	}catch(err)
	{
	     alert("ERROR in insertSlideAction(): "+err.description);
	}
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
			}catch(e)
			{
			  alert("ERROR in SimpleAjaxComponentInsert(): "+e.description);
			}
                   }
    });
}

function simpleAjaxSlideInsert(contenturl,rId, buttonIndex)
{ 
    $.ajax({
          type: "GET",
          url: "search/insert-slide.xqy",
          data: "uri=" + contenturl+"&rid="+ rId,
          success: function(msg){
			try{
			 insertSlideContent(msg,buttonIndex);
			}catch(e)
			{
			  alert("ERROR in SimpleAjaxSlideInsert(): "+e.description);
			}
                   }
    });
}


function insertImage(picuri)
{
       var config = MLA.getConfiguration();
       var fullurl= config.url;
       var picuri = fullurl + "/search/download-support.xqy?uid="+picuri;

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
	}catch(err)
	{
               alert("ERROR in setUndoButton(): "+err.description); 
	}

}

function insertComponentContent(content, other, buttonIndex)
{       //take last JSON string for creation of Shape, for each tag on shape, loop through and add custom parts
	//other is a picture currently, but could include more, hence 'other'
	try{
		var local = MLA.createXMLDOM(content);
		var metaparts = local.getElementsByTagName("dc:metadata");
		var mplength = metaparts.length;

		var jsonPkg = null;
		var jsonString = "";
                var shapeRange = null;
		
		if(mplength > 1)
		{
			var metaXml = metaparts[metaparts.length-1];
			jsonPkg = metaXml.getElementsByTagName("dc:description")[0];
		}
		else
		{
			jsonPkg = metaparts[0].getElementsByTagName("dc:description")[0];
		}


                jsonString = jsonPkg.childNodes[0].nodeValue;
		
		if(jsonString==null || jsonString =="")
		{
			//do nothing
		}
		else
		{
			shapeRange = MLA.jsonParse(jsonString);
			var newShapeName = "";

			if(shapeRange.shape.type=="msoPicture")
			{
				  insertImage(other);
				  setPictureFormat(shapeRange.pictureFormat);
				  var slideIndex=MLA.getSlideIndex();
				  var newShapeName= MLA.getShapeRangeName();
				  var jsonTags = MLA.jsonStringify(shapeRange.tags);
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

function insertSlideContent(content,buttonIndex)
{
    var local = MLA.createXMLDOM(content);
    var metaparts = local.getElementsByTagName("dc:metadata");
    var mplength = metaparts.length;

    var jsonPkg = null;
    var jsonString = "";
    var shapeRange = null;

    for (var i = 0; i < metaparts.length; i++) 
    { 
        //alert(metaparts[i].xml);
        MLA.addCustomXMLPart(metaparts[i].xml);
    }

    setUndoButton(buttonIndex);		
	
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

        if(customPartIds.length > 0 )
	{
	   for (i = 0; i < customPartIds.length; i++)
	   {
               customPartId = customPartIds[i];
	       var customPart = MLA.getCustomXMLPart(customPartId);
               var id = customPart.getElementsByTagName("dc:identifier")[0];
	       //18649719 - from test doc
	       if(id.childNodes[0].nodeValue==ctrlId)
	       {
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
	for(var i = 2;i < meta.childNodes.length; i++)
	{
	        var formID="form-"+i+"-"+controlID;
                var value = $('#'+formID).val();
		meta.childNodes[i].text = value;
	}

        //save edited part
	replaceCustomMetadataPart(metadataPartID, meta);

	$('#ML-Message').show().fadeOut(1500);
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
        if(form.children('div').length)
        {
		form.children('div').remove();
	}
}

function setTagFocus(enteredId)
{

    if( $('#metadata').is(':visible'))  //ONLY DO WHEN TREE EXPOSED, MOVE TO EVENT 
    {
	var tagID = null;
	if(enteredId == null || enteredId == ""){
	        //window.event.cancelBubble=true;
		tagID = window.event.srcElement.id
	}

        //set highlight of selected using class
	$('#treelist').find('a').removeClass("selectedtreectrl");
	$('#'+tagID).addClass("selectedtreectrl");


	//clear metaform in panel
	clearMetadataForm();
	var metaform = $('#metadataForm');

	//need to grab custom piece for metadata section
	var metadataID = getMetadataPartID(tagID);

	if(!(metadataID == null))
	{		
	   var metadata = MLA.getCustomXMLPart(metadataID);
	   var meta = metadata.getElementsByTagName("dc:metadata")[0];

           //start at 2 to skip identifier and first description
	   for(var i = 2;i < meta.childNodes.length; i++)
	   {
		//assumes XML has QName prefix
	        var localname = meta.childNodes[i].nodeName.split(":");
		var formID = "form-"+i+"-"+tagID;
                var child = meta.childNodes[i];	        	
	        var input="";
		var formValue="";

		if(child.childNodes[0] == null)
		{
		        formValue = "";
		}else
		{
			formValue = child.childNodes[0].nodeValue;
		}

		if(localname[1]=="description")
		{
			input = "<textarea cols='40' rows='5' wrap='virtual' id='"+formID+"'>"+
				 formValue +
				"</textarea>";
		}
		else
		{
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

function refreshTagTree()
{
     //didn't have to do this with word as we cleared them with an event

     $('#properties').children('div').remove();

     if($('#treelist').children('li').length)
     {   
	 $('#treelist').children('li').remove();
	 $('#treelist').children('ul').remove();
     }

     var myList = $('#treelist');

     var docTitle = MLA.getPresentationName();

     if( $('#icon-meta-pptxctrl').is('.selectedctrl'))
     {

        var pres_tags = MLA.getPresentationTags();
 
	var iconType = "textIcon";

        if(pres_tags.tags.length==0)
	{
			    myList.append("<li>"+
		            "<a href='#' id='noprestags'>"+
				"<span id='noprestags'>"+
				  "The Presentation has not been tagged."+
                                "</span>"+
			     "</a>"+
			    "</li>");
	}
	else
	{
	  for(var i =0;i<pres_tags.tags.length;i++)
	  {
	    var tag = pres_tags.tags[i];
	    var value = tag.value;
	    if(value=="" || value==null)
	    {
		    //alert("NULL"+value);
	    }else
	    {
	
	    myList.append("<li>"+
		            //"<a href=\"javascript:setNewTagFocus('"+value+"')\" id='"+value+"'>"+
		            "<a href='#' id='"+value+"'>"+
				"<span id='"+value+"'>"+
				     tag.name +
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
	

     }else if( $('#icon-meta-slidectrl').is('.selectedctrl'))
     {
	var slide_index = MLA.getSlideIndex();
        var slide_tags = MLA.getSlideTags(slide_index);
	var iconType = "textIcon";

	if(slide_tags.tags.length==0)
	{
			    myList.append("<li>"+
		            "<a href='#' id='noslidetags'>"+
				"<span id='noslidetags'>"+
				  "The Slide has not been tagged."+
                                "</span>"+
			     "</a>"+
			    "</li>");
	}
        else
	{
	  for(var i =0;i<slide_tags.tags.length;i++)
	  {
	    var tag = slide_tags.tags[i];
	    var value = tag.value;
	    if(value=="" || value==null)
	    {
		    //do nothing
	    }else
	    {
	
	    myList.append("<li>"+
		            //"<a href=\"javascript:setNewTagFocus('"+value+"')\" id='"+value+"'>"+
		            "<a href='#' id='"+value+"'>"+
				//"<span class='"+iconType+"' id='"+value+"'>"+
				"<span id='"+value+"'>"+
				     tag.name +
                                "</span>"+
			     "</a>"+
			    "</li>");

	    }

	    var bref = $('#'+value);

	    bref.bind('click', function() {
                        setTagFocus(); 
                     });
	  }
	}
     }else
     {
	     var shapename="";
	     try{
	         var shapename = MLA.getShapeRangeName();
	     }catch(err)
	     {
		     //donothing
	     }

	     var slideIndex = MLA.getSlideIndex();

	     if(shapename == "" || shapename == null)
	     {
		     var slideIndex = MLA.getSlideIndex();
		     var tagged = checkForComponentTags(slideIndex);

		     if(tagged)
		     {
		           myList.append("<li>"+
		            "<a href='#' id='noslidetags'>"+
				"<span id='noshapeselected'>"+
				  "Slide Includes Tagged Components."+
                                "</span>"+
			     "</a>"+
			    "</li>");
		     }else
		     {
			   myList.append("<li>"+
		            "<a href='#' id='noslidetags'>"+
				"<span id='noshapeselected'>"+
				  "No Components have been Tagged in this Slide."+
                                "</span>"+
			     "</a>"+
			    "</li>");
		     }

		     clearMetadataForm();

	     }
	     else
	     {
		     var shapeRange = MLA.getShapeRangeView(slideIndex, shapename);
		     var iconType =  shapeRange.shape.type;
                     if(shapeRange.tags.length==0)
	             {
			    myList.append("<li>"+
		            "<a href='#' id='noslidetags'>"+
				"<span id='noshapetags'>"+
				  "The Component has not been tagged."+
                                "</span>"+
			     "</a>"+
			    "</li>");
	             }
                     else
	             {
		     
	               for(var i =0;i<shapeRange.tags.length;i++)
	               {
	  	           var tag = shapeRange.tags[i];
	                   var value = tag.value;
	                   if(value=="" || value==null)
	                   {
		                 //alert("NULL"+value);
	                   }else
	                   {
	
	                     myList.append("<li>"+
		                             "<a href='#' id='"+value+"'>"+
				                "<span id='"+value+"'>"+
				                     tag.name +
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
	       "PPT"+currentTime.getTime()+randomNum;
 
    return id;
}

function checkPresentationTags(value)
{
    var pres_tags = MLA.getPresentationTags();

    var same = false;
    for(var i =0;i<pres_tags.tags.length;i++)
    {
        var tag = pres_tags.tags[i];
        if(tag.name.toUpperCase() == value.toUpperCase())
	{
            same = true;
	    break;	
	}
    }

    return same;
}

function checkSlideTags(value)
{
    var slide_index = MLA.getSlideIndex();
    var slide_tags = MLA.getSlideTags(slide_index);

    var same = false;
    for(var i =0;i<slide_tags.tags.length;i++)
    {
	var tag = slide_tags.tags[i];
	if(tag.name.toUpperCase() == value.toUpperCase())
	{
             same = true;
	     break;	
	}
    }

    return same;
}

function checkShapeTags(value)
{
    var slideIndex = MLA.getSlideIndex();
    var shape_name = MLA.getShapeRangeName();
    var shapeRange = MLA.getShapeRangeView(slideIndex, shape_name);
    var shape_tags = shapeRange.tags;
    var same = false;
    for(var i =0;i<shape_tags.length;i++)
    {
	var tag = shape_tags[i];
	if(tag.name.toUpperCase() == value.toUpperCase())
	{
	    same = true;
	    break;	
	}
    }
    return same;
}


function deleteCustomPart(partId)
{
    var customPieceIds = MLA.getCustomXMLPartIds();
    var customPieceId = null;

    if(customPieceIds.length > 0 )
    {
	for (i = 0; i < customPieceIds.length; i++)
	{
            customPieceId = customPieceIds[i];
	    var customPiece = MLA.getCustomXMLPart(customPieceId);
	    //alert(customPiece.xml);
            var id = customPiece.getElementsByTagName("dc:identifier")[0];
	    //18649719 - from test doc
	    if(id.childNodes[0].nodeValue==partId)
	    {
	       MLA.deleteCustomXMLPart(customPieceId);
	    }
	}
    }
}


function deletePresentationTag(tagname, tagvalue)
{
    MLA.deletePresentationTag(tagname);
    deleteCustomPart(tagvalue);
    setPresentationProperties();
}

function deleteSlideTag(slideindex, tagname, tagvalue)
{
    MLA.deleteSlideTag(slideindex, tagname);
    deleteCustomPart(tagvalue);
    setSlideProperties();
}

function deleteShapeTag(slideindex, shapename, tagname, tagvalue)
{
    MLA.deleteShapeTag(slideindex, shapename, tagname);
    deleteCustomPart(tagvalue);
    setShapeProperties();
}


function setPresentationProperties()
{

    var presProps = $('#properties');
    presProps.children('div').remove();
    $('#properties').show();
    $('#noproperties').hide();

    var pres_tags = MLA.getPresentationTags();
    var tagHtml = "";
	
    for(var i =0;i<pres_tags.tags.length;i++)
    {
        var tag = pres_tags.tags[i];
	tagHtml += "<a href=\"javascript:deletePresentationTag('"+tag.name+"','"+tag.value+"')\" id='"+tag.value+"'>"+
		      "<span class='deleteIcon' title='delete tag'><strong><label>"+tag.name+"</label></strong></span>"+
	           "</a>"+
	           "<br/>";

	//delete icon will use tag.value for customxml part,tag.name for tag

    }

    if(tagHtml=="")
    {
	presProps.append("<div id='properties'>Presentation has no Tags</div>");
    }
    else
    {
	presProps.append("<div id='properties'>Presentation has been tagged:<br/><br/>"+tagHtml+"</div>");
    }
	
}

function setSlideProperties()
{
    var slideProps = $('#properties');
    slideProps.children('div').remove();
    $('#properties').show();
    $('#noproperties').hide();
    
    var slide_index = MLA.getSlideIndex();
    var slide_tags = MLA.getSlideTags(slide_index);
    var tagHtml = "";
	
    for(var i =0;i<slide_tags.tags.length;i++)
    {
	var tag = slide_tags.tags[i];
	tagHtml += "<a href=\"javascript:deleteSlideTag('"+slide_index+"','"+tag.name+"','"+tag.value+"')\" id='"+tag.value+"'>"+
 		      "<span class='deleteIcon' title='delete tag'><strong><label>"+tag.name+"</label></strong></span>"+
		   "</a>"+
		   "<br/>";
		//delete icon will use tag.value for customxml part,tag.name for tag
    }

    if(tagHtml=="")
    {
	slideProps.append("<div id='properties'>Slide has no Tags</div>");
    }
    else
    {
	slideProps.append("<div id='properties'>Slide has been tagged:<br/><br/>"+tagHtml+"</div>");
    }
}

function undoInsert(searchType, slideIndex, shapeName)
{
    //var searchType =$("input[@name='search:bst']:checked").val();
    //var slideIndex = MLA.getSlideIndex();
    if(searchType == null || searchType =="")
    {
	alert("You must first insert content before you can undo it");
    }
    else
    {

	if(searchType == "slide")
	{
		var slide_tags = MLA.getSlideTags(slideIndex);
                for(var i =0;i<slide_tags.tags.length;i++)
		{
	            var tag = slide_tags.tags[i];
		    var tagName = tag.name;
	            var tagValue = tag.value;
		    MLA.deleteSlideTag(slideIndex, tagName);
	            deleteCustomPart(tagValue);

		}

	        var slideShapeNames = MLA.getSlideShapeNames(slideIndex);

		for(var j=0;j<slideShapeNames.length;j++)
                {
		      
		      var shapeName = slideShapeNames[j];
		      var shapeRange = MLA.getShapeRangeView(slideIndex, shapeName);
                      var shape_tags = shapeRange.tags;
	              for(var i =0;i<shape_tags.length;i++)
		      {
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
	        for(var i =0;i<shape_tags.length;i++)
		{
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

function checkForComponentTags(slideIndexToCheck, shapeNameToCheck)
{
	var tagged = false;
	var slideIndex = slideIndexToCheck; //MLA.getSlideIndex();

	if(shapeNameToCheck == null || shapeNameToCheck =="")
	{
	    var slideShapeNames = MLA.getSlideShapeNames(slideIndex);

	    for(var j=0;j<slideShapeNames.length;j++)
  	    {
		      
	       var shapeName = slideShapeNames[j];
               var shapeRange = MLA.getShapeRangeView(slideIndex, shapeName);
               var shape_tags = shapeRange.tags;
	       if(shape_tags.length > 0 )
	       {
                  tagged=true;
		  break;
	       }
	    }
	}else
	{
		var shapeRange = MLA.getShapeRangeView(slideIndex, shapeNameToCheck);
		var shape_tags = shapeRange.tags;
                if(shape_tags.length > 0 )
	        {
                    tagged=true;
	        }
	}

	return tagged;

}

function setShapeProperties()
{
    var shapeProps = $('#properties');
    shapeProps.children('div').remove();
    $('#properties').show();
    $('#noproperties').hide();
	
    var shape_name="";

    try{
 	   shape_name = MLA.getShapeRangeName();
    }catch(err)
    {
		//do nothing
    }

    var tagHtml = "";
    if(shape_name == "" || shape_name == null)
    {
	var slideIndex = MLA.getSlideIndex();
        var tagged = checkForComponentTags(slideIndex);

	if(tagged)
	{
          shapeProps.append("<div id='properties'>Slide Includes Tagged Components</div>")
	}
	else
	{
	  shapeProps.append("<div id='properties'>No Components have been Tagged in this Slide</div>");
	}

    }else
    {
	  var slideIndex = MLA.getSlideIndex();
	  var shapeRange = MLA.getShapeRangeView(slideIndex, shape_name);
          var shape_tags = shapeRange.tags;
	  for(var i =0;i<shape_tags.length;i++)
	  {
              var tag = shape_tags[i];
	      tagHtml += "<a href=\"javascript:deleteShapeTag('"+slideIndex+"','"+shape_name+"','"+tag.name+"','"+tag.value+"')\" id='"+tag.value+"'>"+
			    "<span class='deleteIcon' title='delete tag'>&nbsp;</span>"+
			 "</a>"+
			 "<a href=\"javascript:updateComponentJSON('"+slideIndex+"','"+shape_name+"')\">"+
			    "<span class='saveAllIcon' title='save component info'><strong><label>"+tag.name+"</label></strong></span>"+
			 "</a>"+
			 "<br/>";
	  }

	  if(tagHtml=="")
	  {
	      shapeProps.append("<div id='properties'>Component has no Tags</div>");
	  }
	  else
	  {
              shapeProps.append("<div id='properties'>Component has been tagged:<br/><br/>"+tagHtml+"</div>");
	  }
    }
    return false;
}

function refreshPropertiesPanel()
{
    if( $('#icon-pptxctrl').is('.selectedctrl'))
    {
	setPresentationProperties();
    }else if($('#icon-slidectrl').is('.selectedctrl'))
    {
	setSlideProperties();
    }
    else
    {
	setShapeProperties();
    }

}

function getJsonString(shapeRange)
{
    var shapeString = "";

    try{
         shapeString = MLA.jsonStringify(shapeRange);
    }
    catch(err)
    {
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
	   setShapeProperties();
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

function slideSelectionHandler(slideIndex)
{

    if( $('#icon-slidectrl').is('.selectedctrl'))
    { 
        setSlideProperties();
    }
   
    if($('#metadata').is(':visible'))
    {
        refreshTagTree();
	clearMetadataForm();
    }

    if($('#icon-shapectrl').is('.selectedctrl'))
    {
	refreshPropertiesPanel();
    }

}
//END EVENT HANDLERS
