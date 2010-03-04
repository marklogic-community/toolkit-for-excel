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


//BEGIN EVENT HANDLERS
function onEnterHandler(ref)
{
    if( $('#current-doc').is(':visible') ) 
    {
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
      //set properties in Properties Section
      //alert("VISIBLE");
    }
    else 
    {
      //alert("NOT VISIBLE");
    }
}

function onExitHandler(ref)
{
    if( $('#current-doc').is(':visible') ) 
    {
       $('#properties').hide();
       $("#ctrltitle").text("Title: ");
       $("#ctrltag").text("Tag: ");
       $('#lockctrl').attr('checked', false);
       $('#lockcntnt').attr('checked', false);
      //set properties in Properties Section
     // alert("VISIBLE");
    }
    else 
    {
      //alert("NOT VISIBLE");
    }
}

function afterAddHandler(ref)
{
        //here, get parent content control info
	//use title to access map (generated from config.xqy) 
	//and retrieve metadata form to use
	//add part , setting id in custom part to associate
	
	var stringxml = MLA.unescapeXMLCharEntities(generateTemplate(map.get(MLA.getLastAddedControlTitle())));
        var domxml = MLA.createXMLDOM(stringxml);

	//domxml.childNodes[0].childNodes[0];
	var id = domxml.getElementsByTagName("dc:identifier")[0];

	if(id.hasChildNodes())
	{
		alert("HAS CHILDREN");
		id.nodeValue="";
		id.nodeValue=ref.id;
	}
	else
	{
	        //alert("NO CHILDREN");
		var child = id.appendChild(domxml.createTextNode(ref.id));
	}

	MLA.addCustomXMLPart(domxml.xml);
	alert(domxml.xml);

}

function beforeDeleteHandler(ref)
{
	//alert("BEFORE DELETE HANDLER");
	//loop thru custom parts and delete part where id = ref.id
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

function partsTest()
{
	alert("IN TEST");
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

