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


function setControlFocus()
{
	window.event.cancelBubble=true;
        
        //alert(window.event.srcElement.id);
	MLA.setContentControlFocus(window.event.srcElement.id);
	//need to open metadata form in section below tree
	//can we read element names? use these for labels?
	//just open text fields for now.  add value, save to part on keyup
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
     //alert("REFRESHING TREE"+$('#treelist').children('li').length);
     //<li>test</li>
     //<ul><li>test2</li></ul>
     if($('#treelist').children('li').length)
     {   
	 //alert("in the if REMOVING");
	 $('#treelist').children('li').remove();
	 $('#treelist').children('ul').remove();


     }

     var controls = MLA.getSimpleContentControls();
     var control = null;
     var myList = $('#treelist');

    // alert("CONTROLS LENGTH: "+controls.length);
    
     for (i = 0; i < controls.length; i++)
     {
         control=controls[i];
	 var pId = control.parentID;
         var regId = control.id;
	 var iconType = getIconType(control.type);
	 //alert(iconType);

	 if(pId == null || pId.length < 1 )
	 {
	        myList.append("<li>"+
				// "<a href='#' onclick='setControlFocus("+control.id+")'>"+
				 //"<a href='#' id='"+control.id+"'>"+
				 "<a href='#' id='"+regId+"'>"+
				  "<span class='"+iconType+"' id='"+regId+"'>"+
				     control.title +
                                  "</span>"+
				 "</a>"+
			       "</li>");

		var aref = $('#'+regId);
		//alert(regId);
	        aref.bind('click', function() {
                        setControlFocus(); 
			//alert("ID"+regId);
                });

	//	alert("IN THE FIRST IF PARENT UL LENGTH"+ $('#'+regId).parents('ul').length);
	//	alert("IN THE FIRST IF PARENT LI LENGTH"+ $('#'+regId).parents('ul').length);

	/*	var aref = $('#'+control.id).children('a');
			 aref.bind('click', function() {
                    alert('User clicked on id: '+control.id);
                  });
		  */
                //element.attachEvent('onclick',doSomething)
                //$('#foo').bind('click', function() {
                 //alert('User clicked on "foo."');
                //});
		//

	 }else
	 {
		//GET ELEMENT BY ID FOR PARENT, APPEND
		//IF UL ALREADY EXISTS, APPEND LI
		//ELSE APPEND UL, LI


			//alert("PARENT  UL LENGTH: "+ $('#'+pId).parent('ul').length);
			//alert("PARENTS UL LENGTH: "+ $('#'+pId).parents('ul').length);

			var ulLength =  $('#'+pId).parents('ul').length;
			var padding =ulLength * 20;
			//alert("PADDING"+padding);

	    //    if($('#'+pId).parents('ul').length)
	 //	{
	//	alert("IN THE IF");	

			$('#'+pId).parents('ul').eq(0).append("<ul><li>"+
					      "<a href='#' style='padding-left:"+padding+"px' id='"+regId+"'>"+
					      "<span class='"+iconType+"' id='"+regId+"'>"+
					          control.title+
					      "</span>"+
					      "</a>"+
					  "</li></ul>");

			//alert(regId);
		        var bref = $('#'+regId);
	                  bref.bind('click', function() {
			//alert("ID"+regId);
                          setControlFocus(); 
                        });
		
	/*	}
		else
		{	
			alert("IN THE ELSE: "+$('#'+pId).parents('ul').length);

	                $("#"+pId).parents('li').eq(0).append("<ul>"+
					    "<li>"+
					       "<a href='#' style='padding-left:40px'id='"+regId+"'>"+
					        "<span class='"+iconType+"' id='"+regId+"'>"+
  				                   control.title+
						   "</span>"+
					       "</a>"+
					     "</li>"+
					  "</ul>");
			//alert(regId);
		        var cref = $('#'+regId);
	                  cref.bind('click', function() {
		//	alert("ID"+regId);
                         setControlFocus();
                        });

		}
		*/
		

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
		id.nodeValue="";
		id.nodeValue=ref.id;
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

