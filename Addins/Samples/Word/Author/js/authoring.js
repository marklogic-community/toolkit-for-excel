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


var SERVER = "http://localhost:8023/Author";
var BOILERPLATE_URL =  SERVER + "/utils/fetchboilerplate.xqy";

$(document).ready(function() {

       //SET DEFAULTS		
       //by default, current doc selected

       //pro tip: separate with commas , performance?
       //not sure it matters for this app, leaving as is for readability
       $('#metadata').hide();
       $('#search').hide();
       $('#compare').hide();
       $('#properties').hide();
   
       //by default, controls tab selected	
       $('#snippets').hide();

       //by default, textcontrols selected
       $('#imgcontrols').hide();
       $('#calcontrols').hide();
       $('#dropcontrols').hide();
       $('#combocontrols').hide();

       //by default don't show search filtes
       $('#searchfilter').hide();

       //BEGIN top icon tab navigation selection	
       //display current doc tab
       $('a#icon-word').click(function() {

          $('#main').css('overflow', 'hidden');
	  //css
	  $("#icon-metadata").removeClass("fronticon");
	  $("#icon-search").removeClass("fronticon");
	  $("#icon-merge").removeClass("fronticon");
          $("#icon-word").addClass("fronticon");     
   
	  //action
          $('#metadata').hide();
          $('#search').hide();
          $('#compare').hide();
          $('#current-doc').show();

	  $('#docnames').empty();

	  $('#metadataForm').children('div').remove();
   
          return false;
  
       });

       //display metadata icon tab
       $('a#icon-metadata').click(function() {
         
          $('#main').css('overflow', 'hidden');

	  //css
	  $("#icon-word").removeClass("fronticon");
	  $("#icon-search").removeClass("fronticon");
	  $("#icon-merge").removeClass("fronticon");
          $("#icon-metadata").addClass("fronticon");     
   
	  //action 
          $('#current-doc').hide();
          $('#search').hide();
          $('#compare').hide();
          $('#metadata').show();
     
	  $('#docnames').empty();
           
	  refreshControlTree();

          return false;
  
       });

       //display search icon tab
       $('a#icon-search').click(function() {

          $('#main').css('overflow', 'auto');
 	  //css
	  $("#icon-word").removeClass("fronticon");
	  $("#icon-metadata").removeClass("fronticon");
	  $("#icon-merge").removeClass("fronticon");
          $("#icon-search").addClass("fronticon");     
   
	  //action  
          $('#current-doc').hide();
          $('#metadata').hide();
          $('#compare').hide();
          $('#search').show();

	  $('#docnames').empty();

	  $('#metadataForm').children('div').remove();
   
          return false;
  
       });

       //display compare icon tab
       $('a#icon-merge').click(function() {

          $('#main').css('overflow', 'hidden');
 	  //css
	  $("#icon-word").removeClass("fronticon");
	  $("#icon-search").removeClass("fronticon");
	  $("#icon-metadata").removeClass("fronticon");
          $("#icon-merge").addClass("fronticon");     
   
	  //action
          $('#current-doc').hide();
          $('#metadata').hide();
          $('#search').hide();
          $('#compare').show();
   
	  $('#metadataForm').children('div').remove();

          return false;
  
       });
       //END top icon tab navigation selection
       
       //BEGIN control type selection
       //display text control palette
       $('a#icon-textctrl').click(function() {

 	  //css
	  $("#icon-imgctrl").removeClass("selectedctrl");
	  $("#icon-calctrl").removeClass("selectedctrl");
	  $("#icon-dropctrl").removeClass("selectedctrl");
          $("#icon-comboctrl").removeClass("selectedctrl");     
          $("#icon-textctrl").addClass("selectedctrl");     
   
	  //action
          $('#imgcontrols').hide();
          $('#calcontrols').hide();
          $('#dropcontrols').hide();
          $('#combocontrols').hide();
          $('#textcontrols').show();
   
          return false;
  
       });

       //image control palette
       $('a#icon-imgctrl').click(function() {

 	  //css
	  $("#icon-textctrl").removeClass("selectedctrl");
	  $("#icon-calctrl").removeClass("selectedctrl");
	  $("#icon-dropctrl").removeClass("selectedctrl");
          $("#icon-comboctrl").removeClass("selectedctrl");     
          $("#icon-imgctrl").addClass("selectedctrl");     
   
	  //action
          $('#textcontrols').hide();
          $('#calcontrols').hide();
          $('#dropcontrols').hide();
          $('#combocontrols').hide();
          $('#imgcontrols').show();
   
          return false;
  
       });

       //calendar control palette
       $('a#icon-calctrl').click(function() {

 	  //css
	  $("#icon-textctrl").removeClass("selectedctrl");
	  $("#icon-imgctrl").removeClass("selectedctrl");
	  $("#icon-dropctrl").removeClass("selectedctrl");
          $("#icon-comboctrl").removeClass("selectedctrl");     
          $("#icon-calctrl").addClass("selectedctrl");     
   
	  //action
          $('#textcontrols').hide();
          $('#imgcontrols').hide();
          $('#dropcontrols').hide();
          $('#combocontrols').hide();
          $('#calcontrols').show();
   
          return false;
  
       });

       //dropdown control palette
       $('a#icon-dropctrl').click(function() {

 	  //css
	  $("#icon-textctrl").removeClass("selectedctrl");
	  $("#icon-imgctrl").removeClass("selectedctrl");
	  $("#icon-calctrl").removeClass("selectedctrl");
          $("#icon-comboctrl").removeClass("selectedctrl");     
          $("#icon-dropctrl").addClass("selectedctrl");     
   
	  //action
          $('#textcontrols').hide();
          $('#imgcontrols').hide();
          $('#calcontrols').hide();
          $('#combocontrols').hide();
          $('#dropcontrols').show();
   
          return false;
  
       });

       //combobox control palette
       $('a#icon-comboctrl').click(function() {

 	  //css
	  $("#icon-textctrl").removeClass("selectedctrl");
	  $("#icon-imgctrl").removeClass("selectedctrl");
	  $("#icon-calctrl").removeClass("selectedctrl");
          $("#icon-dropctrl").removeClass("selectedctrl");     
          $("#icon-comboctrl").addClass("selectedctrl");     
   
	  //action
          $('#textcontrols').hide();
          $('#imgcontrols').hide();
          $('#calcontrols').hide();
          $('#dropcontrols').hide();
          $('#combocontrols').show();
   
          return false;
  
       });
       //END control type selection 

       //BEGIN current doc controls/snippets selection
       //shows the snippets on clicking the noted link  
       //hides the controls on clicking the noted link
   
        $('a#snippets-show').click(function() {
  
          $("#controltab").removeClass("fronttab");
          $("#snippettab").addClass("fronttab");
          $('#controls').hide();
          $('#snippets').show();
   
          return false;
  
        });

       // hides the snippets on clicking the noted link  
       // shows the controls on clicking the noted link

        $('a#controls-show').click(function() {

          $("#snippettab").removeClass("fronttab");
          $("#controltab").addClass("fronttab");
          $('#controls').show();
          $('#snippets').hide();

          return false;

        });
       //END current doc controls/snippets selection
      
       //Blur ctrlbuttons 
       $('#textcontrols').click(function() {
		      
          $("#buttongroup").li.a.blur();
          return false;
        });

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
	   //alert("startidx="+startidx);
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

	//$('#searchfilter')
	var qry = $('#searchbox').val();
	simpleAjaxSearch(qry,startidx, cbsid);
}

function simpleAjaxSearch(searchval, startidx, cbsid)
{
    var newurl = "";

    if(startidx==0)
	    newurl = "search/search.xqy";
    else
	    newurl = "search/search.xqy?start="+startidx;

    $.ajax({
          type: "GET",
          url: newurl, //"search/search.xqy",
          data: { qry : searchval, params : cbsid },
          success: function(msg){
			try{
                            $('#searchresults').empty();
                            $('#searchresults').append(msg);
                            $('#searchresults').html(msg);
			}catch(e)
			{
			    alert("ERROR"+e.description);
			}
                   }
     });
}

function insertAction(contenturl, contentpath, other)
{ 
	try{
	     simpleAjaxInsert(contenturl,contentpath);
	}catch(err)
	{
	     alert("ERROR in insertAction(): "+err.description);
	}
}

function simpleAjaxInsert(contenturl,contentpath)
{
    $.ajax({
          type: "GET",
          url: "search/insert.xqy",
          data: "uri=" + contenturl+"&path="+ contentpath,
          success: function(msg){
			try{
			  insertContent(msg);
			}catch(e)
			{
			  alert("ERROR in SimpleAjaxInsert(): "+e.description);
			}
                   }
    });
}

var metadataPartArray = new Array();

function insertContent(content)
{       
	try{
		var local = MLA.createXMLDOM(content);
		var pkgxml =  local.getElementsByTagName("pkg:package")[0];
		var metaxml = local.getElementsByTagName("meta")[0];  //NodeList- length, item(idx)
		var metaparts = metaxml.getElementsByTagName("dc:metadata");
	        //alert("PACKAGE XML IS "+pkgxml.xml);
	        //alert("METAPARTS LENGTH IS "+metaparts.length);
	        //alert("METAPART ARAY LENGTH IS "+metadataPartArray.length);

                for (var i = 0; i < metaparts.length; i++) { 
		  //moving reverse here
		  metadataPartArray.push(metaparts.item(i));
                   
                }
		
		MLA.insertWordOpenXML(pkgxml.xml);

	}catch(e)
	{
		//do something here to clear array
		alert("error: "+e.description);
	}
}

//blur selected
function blurSelected(btn_element)
{
	btn_element.blur();
}

//inserts boilerplate
function boilerplateinsert(bp)
{
	//var config = MLA.getConfiguration();
	//alert(config.url);

		$.get(
			BOILERPLATE_URL, 
			{uri: bp},
			function(responseText){
			      MLA.insertWordOpenXML(responseText);
			},
			"text"
		);

}

function lockControl()
{
	var mlacontrolref=MLA.getParentContentControlInfo();
	if(mlacontrolref.lockcontrol=="False")
	{
		MLA.lockContentControl(mlacontrolref.id);

	}else
	{
		MLA.unlockContentControl(mlacontrolref.id);
	}

}

function lockControlContents()
{
	var mlacontrolref=MLA.getParentContentControlInfo();
	if(mlacontrolref.lockcontents=="False")
	{
		MLA.lockContentControlContents(mlacontrolref.id);

	}else
	{
		MLA.unlockContentControlContents(mlacontrolref.id);
	}
}

function getMetadataPartID(ctrlId)
{

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
        //alert("DELETING CUSTOM PIECE");
	MLA.deleteCustomXMLPart(partId);

	//alert("ADDING CUSTOM PIECE");
	MLA.addCustomXMLPart(metadataPart);
}

function setMetadataPartValues()
{
	//get id of currently selected control
	var mlacontrolref=MLA.getParentContentControlInfo();
        var controlID = mlacontrolref.id;	
	//alert("settingMetadataValues(): "+controlID+" "+mlacontrolref.title);

	//get Part ID of Custom XML Part associated with Control
	var metadataPartID = getMetadataPartID(controlID);
	//alert("PART ID: "+metadataPartID);

	//get Custom XML Part associated with Control
	var metadataPart = MLA.getCustomXMLPart(metadataPartID); /*(controlID)*/
	var meta = metadataPart.getElementsByTagName("dc:metadata")[0];
        
	//set form values in Custom XML Part
	for(var i = 1;i < meta.childNodes.length; i++)
	{
	        var formID="form-"+i+"-"+controlID;
                var value = $('#'+formID).val();
		meta.childNodes[i].text = value;
	}

	//alert("FINAL XML" + meta.xml);
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

function setControlFocus(enteredId)
{
      //cancel event bubbling, IE is wacky
      //window.event.cancelBubble=true;

      if( $('#metadata').is(':visible'))  //ONLY DO WHEN TREE EXPOSED, MOVE TO EVENT 
      {	      
        var controlID = null;

	if(enteredId == null || enteredId == ""){
	        window.event.cancelBubble=true;
		controlID = window.event.srcElement.id;
                //check this, do i need, if you click it, its in view
		var destination = $('#'+controlID).offset().top + $("#treeWindow").scrollTop();
		if(!(isScrolledIntoView(controlID)))
		{
		   //at start treeWindow.scrollTop() = 0; treeWindow.height = 200; ctrlId.offset().top = 85;
                   $("#treeWindow").animate({ scrollTop: destination - 85}, 500 );
		}
	
		MLA.setContentControlFocus(controlID);

	}
	else
	{
		controlID = enteredId;

		var destination = $('#'+controlID).offset().top + $("#treeWindow").scrollTop();
		if(!(isScrolledIntoView(controlID)))
		{
                   $("#treeWindow").animate({ scrollTop: destination - 85}, 500 );
		}
	}

        //set highlight of selected using class
	$('#treelist').find('a').removeClass("selectedtreectrl");
	$('#'+controlID).addClass("selectedtreectrl");

	//clear metaform in panel
	var metaform = $('#metadataForm');
        if(metaform.children('div').length)
	{
		metaform.children('div').remove();
	}

	//need to grab custom piece for metadata section
	var metadataID = getMetadataPartID(controlID);
	if(!(metadataID == null))

	{	
	   var metadata = MLA.getCustomXMLPart(metadataID);
	   var meta = metadata.getElementsByTagName("dc:metadata")[0];
	   //check this
	   //var idxml = metadata.getElementsByTagName("dc:identifier")[0];
	
        /*  <div>
              <p><label>Author</label></p>
              <input id="form1" type="text"/> 
              <p>&nbsp; </p>
            </div>
            <div>
              <p><label>Description</label></p>
              <textarea id="form2"/>
              <p>&nbsp; </p>
            </div>
        */

	   //construct metadata panel on page from metadata part
	   for(var i = 1;i < meta.childNodes.length; i++)
	   {
		//assumes XML has QName prefix
	        var localname = meta.childNodes[i].nodeName.split(":");
		var formID = "form-"+i+"-"+controlID;
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
                     setMetadataPartValues(); 
                   }); 
	   }//end of for

	}//end of if
      }

        return false;	

}

function getIconType(ctrlType)
{
	var type = "";
	if(ctrlType=="wdContentControlRichText")
	{
           type="textIcon";
	}else if(ctrlType=="wdContentControlPicture")
	{
           type="imgIcon";
	}else if(ctrlType=="wdContentControlDate")
	{
	   type="calIcon";
	}else if(ctrlType=="wdContentControlDropdownList")
	{
	   type="dropIcon";
	}
	else if(ctrlType=="wdContentControlComboBox")
	{
	   type="comboIcon";
	}

	return type;
}

function refreshControlTree()
{
     if($('#treelist').children('li').length)
     {   
	 $('#treelist').children('li').remove();
	 $('#treelist').children('ul').remove();
     }

     var controls = MLA.getSimpleContentControls();
     //alert("CONTROLS LENGTH"+controls.length);

     var control = null;
     var myList = $('#treelist');

     for (i = 0; i < controls.length; i++)
     {
         control=controls[i];
	 var pId = control.parentID;
         var regId = control.id;
	 var iconType = getIconType(control.type);
	 var title = "";
         (control.title == "" || control.title == null) ?  title = "&nbsp;" : title = control.title;

	 if(pId == null || pId.length < 1 )
	 {
	        myList.append("<li>"+
				 "<a href='#' id='"+regId+"'>"+
				  "<span class='"+iconType+"' id='"+regId+"'>"+
				     title +
                                  "</span>"+
				 "</a>"+
			       "</li>");

		var aref = $('#'+regId);

	        aref.bind('click', function() {
                        setControlFocus(); 
                });

	 }else
	 {
		//GET ELEMENT BY ID FOR PARENT, APPEND
		//IF UL ALREADY EXISTS, APPEND LI
		//ELSE APPEND UL, LI
		var ulLength =  $('#'+pId).parents('ul').length;
		var padding =ulLength * 20;

			$('#'+pId).parents('ul').eq(0).append("<ul><li>"+
					      "<a href='#' style='padding-left:"+padding+"px' id='"+regId+"'>"+
					      "<span class='"+iconType+"' id='"+regId+"'>"+
					          title+
					      "</span>"+
					      "</a>"+
					  "</li></ul>");

		         var bref = $('#'+regId);
	                   bref.bind('click', function() {
                          setControlFocus(); 
                        });
		
	 }
    }
}

function clearControlProperties()
{
    $('#properties').hide();
    $('#noproperties').show();
    $("#ctrltitle").text("");
    $("#ctrltitle").removeClass();
    $("#ctrltag").text("");
    $("#ctrlparent").text("");
    $('#lockctrl').attr('checked', false);
    $('#lockcntnt').attr('checked', false);
}

function myTestFunction()
{
    var myRef = "test";
}

//BEGIN EVENT HANDLERS
function onEnterHandler(ref)
{
       
       var ctrlParent = null;
       if(!(ref.parentID==""))
	      ctrlParent = MLA.getContentControlInfo(ref.parentID);

       //just go ahead and set control properties, thought not necessarily visible

       $("#ctrltitle").text(ref.title);
       $("#ctrltitle").addClass(getIconType(ref.type));
       $("#ctrltag").text("Tag: "+ref.tag);

       if(!(ctrlParent==null))
           $("#ctrlparent").text("Parent: "+ctrlParent.title);

       if(ref.lockcontrol=="True")
       {
	       $('#lockctrl').attr('checked', true);
       }

       if(ref.lockcontents=="True")
       {
	       $('#lockcntnt').attr('checked', true);
       }

       $('#properties').show();
       $('#noproperties').hide();


       setControlFocus(ref.id);
       return false;
}

function onExitHandler(ref)
{
	 clearControlProperties();
}

function afterAddHandler(ref)
{ 
        var proceed = true;	
	//alert("AFTER ADD TITLE: "+ref.title+" ID: "+ ref.id +" TAG: "+ref.tag+" PARENT: "+ref.parentID);

        //when inserting XML with Controls, id is overwritten in Document, but tag remains
	//check here and set tag = id
	//(we set on add with MLA.addContentControl(), but that's not used here if dealing with XML reuse (insert from search))
	var ctrlId = ref.id;
	var ctrlTag = ref.tag;
	try{
		if(!(ctrlId==ctrlTag))
                     MLA.setContentControlTag(ctrlId, ctrlId);	
	}catch(err){
		alert("error: "+e.description);
	}

	//adding metadata part for control
	
	//global array, metadataPartArray, set for metadata coming in for insertable part. see insertContent()
	//if metadataPartArray.length > 0, 
	//   take first item off of array and add as metadata part for this control
	//   (assumes there is metadata for each part, 
	//     in xqy,  we return <dc:metadata/> for no metadata
	//     if dc:metadata.children().length = 0
	//         follow regular metadata process
	//     update array by removing first item (an inserted piece could have several embedded controls)
	//else follow regular metadata process
	
	if(metadataPartArray.length > 0 )
	{
		metadataPartArray.reverse();
		var meta = metadataPartArray.pop();
		metadataPartArray.reverse();

	        //we use dc:identifier with id from ctrl.  
	        //so metadata brought over should have this value updated with latest control id
	        //unique ids for authors are left to other fields, we use this one to link metadata to the ctrl
	        //(important for delete of ctrl, etc)

	        var previd = meta.getElementsByTagName("dc:identifier")[0];
		
		if(previd == null)
		{   
		    //alert("ID IS NULL, No METADATA");
		    proceed=true;
		}
		else 
		{   if(previd.hasChildNodes())
	            {
		      previd.childNodes[0].nodeValue="";
		      previd.childNodes[0].nodeValue=ref.id;
	            }
	            else
	            {
		      var child = previd.appendChild(meta.createTextNode(ref.id));
	            }

                    MLA.addCustomXMLPart(meta.xml);
		    proceed = false;
		}

	}
	
	//assuming the control was added from the palette on the pane:
	//use title to access map (generated from config.xqy) 
	//and retrieve metadata form to use
	//add part , setting id in custom part to associate
	//possible to move a control, this creates delete, then add event using same id
        if(proceed)
	{
	   
	   //have the server be the broker for the metadata templates
	   //logged RFE 
	   
           var stringxml = MLA.unescapeXMLCharEntities(generateTemplate(map.get(MLA.getLastAddedControlTitle())));
           var domxml = MLA.createXMLDOM(stringxml);

	   var id = domxml.getElementsByTagName("dc:identifier")[0];

	   if(id.hasChildNodes())
	   {
		id.childNodes[0].nodeValue="";
		id.childNodes[0].nodeValue=ref.id;
	   } 
	   else
	   {
	        var child = id.appendChild(domxml.createTextNode(ref.id));
	   }

	   MLA.addCustomXMLPart(domxml.xml);
	}

	return false;

}

function beforeDeleteHandler(ref)
{
	//alert("BEFORE DELETE HANDLER: "+ref.id);
	//loop thru custom parts and delete part where id = ref.id
	//person can move controls, so need to save values if they exist, and use if new piece added with same id
	
	//remove node from TreeView
	//check, may need to remove parent
	if( $('#metadata').is(':visible') )
	{
	    var node = $('#'+ref.id).remove();
	}

	var customPieceIds = MLA.getCustomXMLPartIds();
        var customPieceId = null;

        if(customPieceIds.length > 0 )
	{
	   for (i = 0; i < customPieceIds.length; i++)
	   {
               customPieceId = customPieceIds[i];
	       var customPiece = MLA.getCustomXMLPart(customPieceId);
               var id = customPiece.getElementsByTagName("dc:identifier")[0];
	       //18649719 - from test doc
	       if(id.childNodes[0].nodeValue==ref.id)
	       {
			     //alert("===EQUAL"+id.xml);
			     MLA.deleteCustomXMLPart(customPieceId);
	       }
	       //alert(customPiece.xml);
	    }
	}
}

function siteChanged(selectedOption, startidx)
{
	if(startidx==null)
	{
	   //alert("startidx="+startidx);
	   startidx = 0;
	}

	//var selectedOption = $('#sites').val();
	simpleAjaxMetadataSearch(selectedOption, startidx);
        //alert("HERE: "+selectedOption + " VAL: "+document.getElementById('select0').innerHTML);
}

function simpleAjaxMetadataSearch(searchval, startidx)
{
    var newurl = "";


    if(startidx==0)
	    newurl = "search/metadata-search.xqy";
    else
	    newurl = "search/metadata-search.xqy?start="+startidx;


    $.ajax({
          type: "GET",
          url: newurl, 
          data: { qry : searchval },
          success: function(msg){
			try{
			     //alert("MESSAGE IS: "+msg+ "  "+msg.length);
                             $('#docnames').empty();
                             $('#docnames').append(msg);
                             $('#docnames').html(msg);
			}catch(e)
			{
			alert("ERROR"+e.description);
			}
	                     //alert( "Data Saved: " + msg );
                   }
     });
}

function compareSearchAction(startidx)
{
	if(startidx==null)
	{
	   //alert("startidx="+startidx);
	   startidx = 0;
	}

	var qry = document.getElementById('select0').innerHTML;
	alert(qry, startidx);
	//simpleAjaxMetadataSearch(qry,startidx);
}

function mergeDocuments(docuri)
{
	var selectedOption = docuri;
	if(selectedOption == null)
	{
          //do nothing
	  alert("You must first select a document to merge.");
	}else
	{
	  simpleAjaxDocRetrieve(selectedOption);
	  
	}
}

function simpleAjaxDocRetrieve(doclocation)
{
    var newurl = "search/document-retrieve.xqy";

    $.ajax({
          type: "GET",
          url: newurl, 
          //data: { qry : searchval, params : cbsid },
          data: { qry : doclocation },
          success: function(opc_xml){
			try{
			     MLA.mergeWithActiveDocument(opc_xml);
			}catch(e)
			{
			alert("ERROR"+e.description);
			}
                   }
     });
}

//menu functions
function setSelected(index)
{
	var v_id="select"+index;
	var elem = document.getElementById(v_id);
        var searchparam = elem.name;	
	var orig = document.getElementById('select0');
	orig.innerHTML=elem.innerHTML;
	siteChanged(searchparam);
	//alert(orig.outerHTML);
}

function displayLayer(layer)
{
	var myLayer = document.getElementById(layer);
	if(myLayer.style.display=="none" || myLayer.style.display==""){
		myLayer.style.display="block";
	} else {
		myLayer.style.display="none";
	}
}


//END EVENT HANDLERS

