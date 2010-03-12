$(document).ready(function() {
   
       //by default, current doc selected
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
      
       //BEGIN top icon tab navigation selection	
       //display current doc tab
       $('a#icon-word').click(function() {

	  //location.reload();
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
   
          return false;
  
       });

       //display metadata icon tab
       $('a#icon-metadata').click(function() {

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
     
	  refreshControlTree(); 
          return false;
  
       });

       //display search icon tab
       $('a#icon-search').click(function() {

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
   
          return false;
  
       });

       //display compare icon tab
       $('a#icon-merge').click(function() {

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
       // shows the snippets on clicking the noted link  
       // hides the controls on clicking the noted link
   
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
       

       //Blur ctrlbuttons ??
       $('#textcontrols').click(function() {
		      
          $("#buttongroup").li.a.blur();
          return false;
        }) 
 
});

//window.event.cancelBubble = true;

function blurSelected(btn_element)
{
	btn_element.blur();
}

/*maybe change sig to (url, bp) */
//potentially need another function call in the case of controls, a trigger to insure all controls have a custom part associated
function boilerplateinsert(bp)
{
	       //alert("HERE "+bp);
	
		$.get(
			"http://localhost:8023/KA/utils/fetchboilerplate.xqy",  /*should we place in config somewhere, maybe pass onclick*/
			{uri: bp},
			function(responseText){
			      //alert("RESPONSE: "+responseText);
			      MLA.insertWordOpenXML(responseText);
				//$("#result").html(responseText);
			      //validateControlsFunction here. Check all controls, add metadata where needed.
			},
			"text"
		);

               //alert("FINIS "+bp)
	     
}

function lockControl()
{
	var mlacontrolref=MLA.getParentContentControlInfo();
	if(mlacontrolref.lockcontrol=="False")
	{
		MLA.lockContentControl(mlacontrolref.id);
		//alert("LOCKING CONTROL");

	}else
	{
		MLA.unlockContentControl(mlacontrolref.id);
		//alert("UNLOCKING CONTROL");

	}

	//alert("ENTER ---> message "+mlacontrolref.id+" tag "+mlacontrolref.tag+" title"+mlacontrolref.title+" type "+mlacontrolref.type + "lockcontrol"+ mlacontrolref.lockcontrol + " lockcontents"+ mlacontrolref.lockcontents +" parentTag "+mlacontrolref.parentTag + " parentID: "+ mlacontrolref.parentID);
}

function lockControlContents()
{
	var mlacontrolref=MLA.getParentContentControlInfo();
	if(mlacontrolref.lockcontents=="False")
	{
		MLA.lockContentControlContents(mlacontrolref.id);
		//alert("LOCKING CONTROL CONTENTS");

	}else
	{
		MLA.unlockContentControlContents(mlacontrolref.id);
		//alert("UNLOCKING CONTROL CONTENTS");

	}
}

function partsTest()
{
	alert("IN TEST");
	var control = null;
	var controls = MLA.getSimpleContentControls();

        for (i = 0; i < controls.length; i++)
	{
		control=controls[i];
		alert("ID: "+control.id);
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

function getMetadataPartForControl(customPartId)
{
        var metadataPart =  MLA.getCustomXMLPart(customPartId);
	return metadataPart;

	/*
	var customPieceIds = MLA.getCustomXMLPartIds();
        var customPieceId = null;

	var metadataPart = null;

	//alert(customPieceIds.length);
        if(customPieceIds.length > 0 )
	{
	   for (i = 0; i < customPieceIds.length; i++)
	   {
               customPieceId = customPieceIds[i];
	       var customPiece = MLA.getCustomXMLPart(customPieceId);
               var id = customPiece.getElementsByTagName("dc:identifier")[0];
	       //18649719 - from test doc
	       if(id.childNodes[0].nodeValue==ctrlId)
	       {
			     
			      metadataPart = customPiece;
			      // MLA.deleteCustomXMLPart(customPieceId);
	       }

	       //alert(customPiece.xml);
	    }
	}

	return metadataPart;
	*/
}

//can i reuse based on enter event?
//add id param, if null, then use window.event.cancelBubble, else id (?) - should work
//
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
	var metadataPart = getMetadataPartForControl(metadataPartID); /*(controlID)*/
	var meta = metadataPart.getElementsByTagName("dc:metadata")[0];

	for(var i = 1;i < meta.childNodes.length; i++)
	{
	        var formID="form-"+i+"-"+controlID;
                var value = $('#'+formID).val();
		meta.childNodes[i].text = value;
	}

	//alert("FINAL XML" + meta.xml);

	replaceCustomMetadataPart(metadataPartID, meta);


}

function setControlFocus(enteredId)
{
	//cancel event bubbling, IE is wacky
	//window.event.cancelBubble=true;
         
        var controlID = null;

	if(enteredId == null || enteredId == ""){
	         window.event.cancelBubble=true;
		 controlID = window.event.srcElement.id;
		 MLA.setContentControlFocus(controlID);
	}
	else
	{
		controlID = enteredId;
		var destination = $('#'+controlID).offset().top;
                $("#treeWindow").animate({ scrollTop: destination-60}, 500 );
	}

        //set highlight of selected using class
	$('#treelist').find('a').removeClass("selectedtreectrl");
	$('#'+controlID).addClass("selectedtreectrl");

   //var destination = $('#'+controlID).offset().top;
   //$("#treeWindow").animate({ scrollTop: destination-20}, 500 );
   //return false;
	

	//set focus in document on selected control
	//MLA.setContentControlFocus(controlID);

	//need to grab custom piece for metadata section
	var metadataID = getMetadataPartID(controlID);
	//alert("metadatatID" +metadataID);
	var metadata = getMetadataPartForControl(metadataID);

	var meta = metadata.getElementsByTagName("dc:metadata")[0];


	//alert("META: "+meta.xml);
	var idxml = metadata.getElementsByTagName("dc:identifier")[0];
	
        //alert("ID: "+idxml.xml + idxml.childNodes[0].nodeValue);


	//addMetadata to pane here. set fields based on child elements?
	//assumption id is childNodes[0]
	//
   /*<div>
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
        var metaform = $('#metadataForm');
        if(metaform.children('div').length)
	{
		metaform.children('div').remove();
	}

	//construct metadata panel on page
	for(var i = 1;i < meta.childNodes.length; i++)
	{
	        var localname = meta.childNodes[i].nodeName.split(":");
		var formID = "form-"+i+"-"+controlID;
		//var formID = "form-"+controlID;
                var child = meta.childNodes[i];	        	
		//alert("LOCALNAME: "+localname[1]+"  "+child.nodeName);
	        var input="";
		var formValue="";

		if(child.childNodes[0] == null)
		{
			//alert("IN THE IF for FORMVALUE");
		        formValue = "";
		}else
		{
                       
			formValue = child.childNodes[0].nodeValue;
			//alert("FORM VALUE IN ELSE"+formValue);
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
		  //alert(formID);

	}
	

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
    $("#ctrltitle").text("Title: ");
    $("#ctrltag").text("Tag: ");
    $('#lockctrl').attr('checked', false);
    $('#lockcntnt').attr('checked', false);
}

//BEGIN EVENT HANDLERS
function onEnterHandler(ref)
{
       //just go ahead and set control properties, thought not necessarily visible
       $("#ctrltitle").text("Title: "+ref.title);
       $("#ctrltag").text("Tag: "+ref.tag);
       if(ref.lockcontrol=="True")
       {
	       $('#lockctrl').attr('checked', true);
       }

       if(ref.lockcontents=="True")
       {
	       $('#lockcntnt').attr('checked', true);
       }

       $('#properties').show();

       //alert("ID: "+ref.id);
       //$('#treelist').find('a').removeClass("selectedtreectrl");
       //$('#'+ref.id).addClass("selectedtreectrl");
       setControlFocus(ref.id);


}

function onExitHandler(ref)
{
	 clearControlProperties();
}

function afterAddHandler(ref)
{
	//alert("AFTER ADD: "+ref.id +"PARENT: "+ref.parentID);
        //here, get parent content control info
	//use title to access map (generated from config.xqy) 
	//and retrieve metadata form to use
	//add part , setting id in custom part to associate
	//
	//possible to move a control, this creates delete, then add event using same id
	//if id same, need to add back original metadata values
	//no title, so only refresh if pane is visible
        //if( $('#metadata').is(':visible') )
	//{	
        //   refreshControlTree();
        //}

        var stringxml = MLA.unescapeXMLCharEntities(generateTemplate(map.get(MLA.getLastAddedControlTitle())));
        var domxml = MLA.createXMLDOM(stringxml);

	//domxml.childNodes[0].childNodes[0];
	var id = domxml.getElementsByTagName("dc:identifier")[0];

	if(id.hasChildNodes())
	{
		//alert("HAS CHILDREN");
		id.childNodes[0].nodeValue="";
		id.childNodes[0].nodeValue=ref.id;
	}
	else
	{
	        //alert("NO CHILDREN");
		var child = id.appendChild(domxml.createTextNode(ref.id));
	}

	MLA.addCustomXMLPart(domxml.xml);
	//alert("ADDING" + domxml.xml);

}

function beforeDeleteHandler(ref)
{
	//alert("BEFORE DELETE HANDLER: "+ref.id);
	//loop thru custom parts and delete part where id = ref.id
	//person can move controls, so need to save values if they exist, and use if new piece added with same id
	
	//remove node from TreeView
	if( $('#metadata').is(':visible') )
	{
	    var node = $('#'+ref.id).remove();
	}

	var customPieceIds = MLA.getCustomXMLPartIds();
        var customPieceId = null;

	//alert(customPieceIds.length);
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

//END EVENT HANDLERS
//
//BEGIN METADATA MAPPING
//var myparams;
//var map = new MetadataMap();

//this to be generated from config
/*
function MetadataMap()
{
   myparams = new Array();
   myparams["Section"] = "2";
   myparams["Annex"] = "2";

}
*/
/*MetadataMap.prototype.get = function(key)
{
	return myparams[key];
}
*/
//15 DC elements
//title
//creator
//subject
//description
//publisher
//contributor
//date
//type
//format
//identifier
//source
//language
//relation
//coverage
//rights

//this to be generated from config 
/*
function generateTemplate(metaid) {
    if(metaid == "1")
    {
    var v_template = "<dc:metadata "
		   + "xmlns:dc='http://purl.org/dc/elements/1.1/'>"
		     + "<dc:identifier>" 
		     + "</dc:identifier>" 
		     + "<dc:title>" 
		     +  "</dc:title>"
		     + "<dc:subject>"  
		     + "</dc:subject>" 
		     + "<dc:publisher>"  
		     + "</dc:publisher>"
		     + "<dc:identifier>"  
		     + "</dc:identifier>"
		   + "</dc:metadata>";	
    }else if(metaid == "2"){
             var v_template = "<dc:metadata "
		   + "xmlns:dc='http://purl.org/dc/elements/1.1/'>"
		     + "<dc:identifier>" 
		     + "</dc:identifier>" 
		     + "<dc:contributor>" 
		     + "</dc:contributor>"
		     + "<dc:description>"  
		     + "</dc:description>" 
		   + "</dc:metadata>";	
    }else
    {
        var v_template = "<dc:metadata "
		   + "xmlns:dc='http://purl.org/dc/elements/1.1/'>"
		     + "<dc:identifier>" 
		     + "</dc:identifier>" 
		     + "<dc:contributor>" 
		     + "</dc:contributor>"
		     + "<dc:relation>"  
		     + "</dc:relation>" 
		   + "</dc:metadata>";
    }

   return v_template;
}
*/
/*
function generateTemplate(metaid){if(metaid=='1'){ var v_template='<dc:metadata xmlns:config="http://marklogic.com/config" xmlns:dc="http://purl.org/dc/elements/1.1/"> <dc:identifier/> <dc:title/> <dc:subject/> <dc:publisher/> <dc:description/> </dc:metadata>';}else if(metaid=='2'){ var v_template='<dc:metadata xmlns:config="http://marklogic.com/config" xmlns:dc="http://purl.org/dc/elements/1.1/"> <dc:identifier/> <dc:contributor/> <dc:source/> </dc:metadata>';}else{var v_template='<dc:metadata xmlns:config="http://marklogic.com/config" xmlns:dc="http://purl.org/dc/elements/1.1/"> <dc:identifier/> <dc:contributor/> <dc:relation/> </dc:metadata>';}return v_template;}
*/
		

//END METADATA MAPPING

