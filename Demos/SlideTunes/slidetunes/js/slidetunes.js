var RESTSERVER = "http://localhost:8060";
var APPSERVER = "http://localhost:8030";

var PRESODIR = "/paven/";
var PLAYLISTDIR = "/gallery/";

var SEARCHURI = RESTSERVER +"/office";
var PRESOURI = RESTSERVER+"/office/presentations";
var PLAYLISTURI = RESTSERVER+"/playlists";

var GETLIST = APPSERVER + "/slidetunes/xquery/get.xqy";
var GETSLDS = APPSERVER + "/slidetunes/xquery/get-slides.xqy";

var SAVEPLAYLIST = APPSERVER + "/slidetunes/xquery/save-playlist.xqy";
var DELETEPLAYLIST = APPSERVER + "/slidetunes/xquery/delete-playlist.xqy";
var EXPORTPLAYLIST = APPSERVER + "/slidetunes/xquery/export-playlist.xqy";
var EXISTS = APPSERVER +  "/slidetunes/xquery/exists.xqy";
var OPENPPTX = APPSERVER +  "/slidetunes/xquery/open-binary.xqy";

$(document).ready(function() {

		$(document).bind("contextmenu",function(e){
                    return false;
                });
	 
		$('#deck-playlist ul').hoverscroll({
			width:		"100%",        // Width of the list
			height:		47         // Height of the list
		});
		resizeDeckPlaylist();
		
		// search results are connected... meaning I can drag TO the playlist
		$( "#deck-search-results ul" ).sortable({
			opacity: 0.6,
			connectWith: "#deck-playlist ul"
		}).disableSelection();				
		
		// connect deck viewer to the playlist
		//$( "#deck-viewer-list ul" ).sortable({
		//	opacity: 0.6,
		//	connectWith: "#deck-playlist ul"
		//}).disableSelection();		
		
		$( "#deck-playlist ul" ).sortable({
			opacity: 0.6,
			over: function(event, ui) { 
				resizeDeckPlaylist();
			},
			receive: function(event, ui) { 
				resizeDeckPlaylist();
			},
			update: function(event,ui){
		         	updatePlaylist();
			}
		}).disableSelection();		

		$('.library-deck-btn').live('click',function(ev) {
			var deckContents = '#deck-' + ($(this).parent().attr('id').replace('header-',''));
			$(deckContents).toggle();
			if ($(this).hasClass('closed'))
				$(this).removeClass('closed');
			else
				$(this).addClass('closed');			
		});
		$(window).resize(function() {
			$('#deck-search-results').height($(window).height() - ($('#header').outerHeight() + $('#deck-header').outerHeight() + $('#deck-playlist').outerHeight() + 22));		
		});
		$('#deck-search-results').height($(window).height() - ($('#header').outerHeight() + $('#deck-header').outerHeight() + $('#deck-playlist').outerHeight() + 22));		
		
		$('#deck-viewer-slide-previous').click(function() {
			$('#deck-viewer-list ul li').each(function() {
				if ($(this).hasClass('lg')) {
					$('#deck-viewer-list ul li').removeClass('med');
					if ($(this).prev().prev().length) {
						$(this).prev().prev().addClass('med');
						$(this).prev().addClass('lg');
					}
					$(this).removeClass('lg');
					$(this).addClass('med');
					return false; 
				}
			});
		});
		$('#deck-viewer-slide-next').click(function() {
			$('#deck-viewer-list ul li').each(function() {
				if ($(this).hasClass('lg')) {
					$('#deck-viewer-list ul li').removeClass('med');
					if ($(this).next().next().length) {
						$(this).next().next().addClass('med');
						$(this).next().addClass('lg');
					}
					$(this).removeClass('lg');
					$(this).addClass('med');
					return false; 
				}
			});
		});
		
		$('#deck-lists ul li').live('click',function() {
			$('#deck-viewer').show();
			var myOffset = $(this).outerWidth() + ' -188';
			$('#deck-viewer').position({
                    of: $(this),
                    my: 'left top',
                    at: 'left bottom',
                    offset: myOffset,        
                    collision: "none none"
                });

		$('#deck-viewer').focus();
		});
		
		$('#header, #wrapper').click(function() {
			$('#deck-viewer').hide();
		});

		populateLibraryListing(PRESOURI+PRESODIR, "presentations", randomId());
                populateLibraryListing(PLAYLISTURI+PLAYLISTDIR,"playlists", randomId());

		
	});

resizeDeckPlaylist = function(){
		var size = 0;
		$('#deck-playlist ul').children().each(function() {
			if  (!$(this).hasClass('item'))
				$(this).addClass('item');
			
			size += $(this).width() + parseInt($(this).css('padding-left')) + parseInt($(this).css('padding-right'))
				+ parseInt($(this).css('margin-left')) + parseInt($(this).css('margin-right'));					
		});
		// Apply computed width to listcontainer
		$('#deck-playlist ul').width(size + 20);


}

checkEventElement = function(e){
    var e=e ? e : window.event;
    var event_element=e.target? e.target : e.srcElement;
    return event_element;
}
	
populateLibraryListing = function(uri, destination, rand){
    var presos = simpleAjaxFetchPresentationList(uri, destination, rand);
	
}

simpleAjaxFetchPresentationList = function(uri, destination, rand)
{
    var newurl = GETLIST+"?"+rand;

    $.ajax({
	type: "GET",
	url: newurl,
	data: { geturi : uri },
	success: function(msg){
	  try{

               

	  if(destination == "presentations"){
	     updateLibPresentationList(msg);
	  }
	  else{
	     updateLibPlaylistList(msg);
	  }
	  }catch(e){
	      alert("ERROR"+e.description);
	  }
	}

    });

}


updateLibPresentationList = function(msg){

    try{
        var local =  MLA.createXMLDOM(msg);
        var presos = local.getElementsByTagName("presentation");

	if($('#deck-mydecks').children('li').length){   
	    $('#deck-mydecks').children('li').remove();
     	}

	var presoList = $('#deck-mydecks');

        for( var i = 0; i < presos.length; i++) {
            var pres =  presos[i].childNodes[0].nodeValue;
	    var rId = "pid"+i;

	                                              /*<div id="test">
							<div class="deck-thumb">
							  <img src="images/slide-placeholder-sm.png" />
							</div>
							<div class="deck-details">
								<p class="name">Deck # 1</p>		
								<p class="date">mm/dd/yyyy</p>		
							</div>	
                                                      </div>	*/
	    presoList.append("<li>"+
		               //"<div id='"+rId+"'>"+
		               "<div>"+
			           "<div class='deck-thumb'>"+
					"<img src='images/slide-placeholder-sm.png' />"+
				   "</div>"+
				   "<div class='deck-details'>"+
			        //      "<a href='javascript:presoAction("+'"'+pres+'"'+");'>"+
			                "<p id='"+rId+"'class='name'>"+pres+"</p>"+
				        "<p class='date'>mm/dd/yyyy</p>"+
		                   "</div>"+		     
				"</div>"+
			     "</li>");

	    var aref = $('#'+rId);

	    aref.bind('mousedown', function(e) {
                        //setControlFocus(window.event.srcElement.id);
                e.preventDefault();
	
	         //check for right click
	        if( e.button == 2 ){ 
                    var event_element=checkEventElement(e);
                    //alert(event_element.tagName + event_element.id + event_element.childNodes[0].nodeValue);
                    setContextMenu(event_element.id);
                    return false; 
                } 
            });
	    

        }

    }catch(e){
        alert("ERROR"+e.description);
    }
}

updateLibPlaylistList =  function(msg){
    try{
        var local =  MLA.createXMLDOM(msg);
        var pls = local.getElementsByTagName("playlist");

        if($('#deck-myplaylists').children('li').length){   
	    $('#deck-myplaylists').children('li').remove();
        }
        var plList = $('#deck-myplaylists');

        for (var i = 0; i < pls.length; i++) {
            var pl = pls[i].childNodes[0].nodeValue; 
	    var rId = "plid"+i;

	    plList.append("<li>"+
		               //"<div id='"+rId+"'>"+
		               "<div>"+
			           "<div class='deck-thumb'>"+
					"<img src='images/slide-placeholder-sm.png' />"+
				   "</div>"+
				   "<div class='deck-details'>"+
			        //      "<a href='javascript:presoAction("+'"'+pres+'"'+");'>"+
			                "<p id='"+rId+"'class='name'>"+pl+"</p>"+
				        "<p class='date'>mm/dd/yyyy</p>"+
		                   "</div>"+		     
				"</div>"+
			     "</li>");

         var aref = $('#'+rId);

	 aref.bind('mousedown', function(e) {
                       
         e.preventDefault();
	 //check for right click
	 if( e.button == 2 ){ 
                var event_element=checkEventElement(e);
		e.preventDefault();
                //alert(event_element.tagName + event_element.id + event_element.childNodes[0].nodeValue);
                setContextMenu(event_element.id);
                return false; 
              } 
            });

        }

    }catch(e){
        alert("ERROR"+e.description);
    }
}

setContextMenu = function(rId)
{
    var aref = $('#'+rId)

    var pos = aref.offset();  
    var width = aref.width();
    //show the menu directly over the placeholder
    $(".vmenu").css( { "left": (pos.left) + "px", "top":pos.top + "px" } );
    $(".vmenu").show();
 
    //need to unbind, or we keep binding and end up with multiple 
    $('.vmenu .first_li').bind('click',function() {
	plAction(aref.text());
        $('.vmenu .first_li').unbind('click');
        $('.vmenu').hide();
    });

    $('.vmenu .second_li').bind('click',function() {
	presoAction(aref.text());
        $('.vmenu .second_li').unbind('click');
   	$('.vmenu').hide();
    });

    $('.vmenu .third_li').bind('click',function() {
	deleteAction(aref.text());
        $('.vmenu .third_li').unbind('click');
   	$('.vmenu').hide();
    });

 
    $(".first_li span").hover(function () {
        $(this).css({backgroundColor : '#E0EDFE' , cursor : 'pointer'})
    },
    function () {
	$(this).css('background-color' , '#fff' );
    });

    $(".second_li span").hover(function () {
        $(this).css({backgroundColor : '#E0EDFE' , cursor : 'pointer'})
    },
    function () {
	$(this).css('background-color' , '#fff' );
    });

     $(".third_li span").hover(function () {
        $(this).css({backgroundColor : '#E0EDFE' , cursor : 'pointer'})
    },
    function () {
	$(this).css('background-color' , '#fff' );
    });
}

presoAction = function(presentation){
    var serveruri = PRESOURI;
    var slideuri = presentation + "/slides";
    simpleAjaxFetchImages(serveruri, slideuri, "workspace");
}

plAction = function(playlist){

    clearPlaylistExportLink();
    $(".plname").text(playlist);
    var serveruri = PRESOURI;
    var slideuri = PLAYLISTURI+playlist;
    simpleAjaxFetchImages(serveruri, slideuri, "playlists");
}

deleteAction = function(playlist){
    deletePlaylist(PLAYLISTURI+playlist);  
}

simpleAjaxFetchImages =function(serveruri, slideuri, destination){
   
    var newurl = GETSLDS;
   //alert("ServerURI: "+serveruri+" SlideURI: "+slideuri+" NewURL: "+newurl); 

    $.ajax({
	type: "GET",
	url: newurl,
	data: { srvuri : serveruri, slduri: slideuri, dest: destination },
	success: function(msg){

	  try{

	    if(destination == "workspace"){
	       updateWorkspaceImages(msg);
	    }
	    else{
	       updatePlaylistImages(msg);
	    }
	  }catch(e){
	      alert("ERROR"+e.description);
	  }
	}
    });
}

updateWorkspaceImages = function(msg){
    if($('#deck-search-results').children('ul').length){   
        $('#deck-search-results').children('ul').remove();
    }
    var plList = $('#deck-search-results');
    plList.html(msg);

    $( "#deck-search-results ul" ).sortable({
			opacity: 0.6,
			connectWith: "#deck-playlist ul"
    }).disableSelection();	
}

updatePlaylistImages = function(msg){

    if($('#deck-playlist').children('ul').length){   
         $('#deck-playlist').children('ul').remove();
    }
    var plList = $('#deck-playlist');
    plList.html(msg);

    $('#deck-playlist ul').hoverscroll({
	 width:  "100%",        // Width of the list
	 height:  47         // Height of the list
    });

    $( "#deck-playlist ul" ).sortable({
			opacity: 0.6,
			over: function(event, ui) { 
				resizeDeckPlaylist();
			},
			receive: function(event, ui) { 
				resizeDeckPlaylist();
			},
			update: function(event,ui){
		         	updatePlaylist();
		       	}
    }).disableSelection();		
    
}

updatePlaylist = function(){
    //went to listcontainer class as #deck-playlist has two other children divs before the ul
    //var srcAttrs = $('#deck-playlist').children('ul').children('li').children('span').children('img');
    $(".dummy").remove();

    var plName = $(".plname").text();

    var srcAttrs = $('.listcontainer').children('ul').children('li').children('span').children('img');

    var ACTIVE_PLAYLIST="<playlist><slides>";

    var idxLength = PRESOURI.length;
    srcAttrs.each( function()
 	           {

		      try{
		         var url =  $(this).attr('src');
			 var single = $(this).parent('span').attr('id');
			 
			 ACTIVE_PLAYLIST+="<slide>"+
			                     "<image>"+
					        url.substring(idxLength)+
					     "</image>"+
					     "<single>"+
					         single+
			                     "</single>"+
					   "</slide>";

		      }catch(e){
		         alert("ERROR"+e.description);
	              }

	           });
	

    ACTIVE_PLAYLIST+="</slides></playlist>";
    savePlaylist(plName, ACTIVE_PLAYLIST);
}	

addPlaylist = function(tempname){

    var playlistname = PLAYLISTDIR+tempname+".xml";
    //setName on playlist
    $(".plname").text(playlistname);

    //clear whats in playlist, add empty li so we can add to it
    if($('#deck-playlist').children('ul').length){   
         $('#deck-playlist').children('ul').remove();
    }

    var plList = $('#deck-playlist');
    plList.html("<ul class='connect'><li class='dummy'>Add a Slide Here</li></ul>");

    $('#deck-playlist ul').hoverscroll({
	 width:  "100%",        // Width of the list
	 height:  47         // Height of the list
    });

    $( "#deck-playlist ul" ).sortable({
			opacity: 0.6,
			over: function(event, ui) { 
				resizeDeckPlaylist();
			},
			receive: function(event, ui) { 
				resizeDeckPlaylist();
			},
			update: function(event,ui){
		         	updatePlaylist();
		       	}
    }).disableSelection();		
        
       
     //save empty playlist doc to ML?
     var PLAYLISTDOC = "<playlist><slides></slides></playlist>";
     savePlaylist(playlistname, PLAYLISTDOC);
     
     //IE caches here, need a rand?
     //its asynchronous, need to wait for savePlaylist to return
}

function randomId()
{
    var currentTime = new Date();	
    var randomNum = Math.floor(Math.random()*50000);
    var id =   // currentTime.getHours()+":" +
   	       // currentTime.getMinutes() + ":" +
	       // currentTime.getSeconds() + ":" +
	       "SL"+currentTime.getTime()+randomNum;
 
    return id;
}


function savePlaylist(playlistName, galleryXml)
{ 
    //TODO clean up this and save-playlist.xqy, its not using PLAYLIST URI, couldn't PUT XML?  only text/binary?  so doing direct xdmp:document-insert
    var newurl = SAVEPLAYLIST; //"/xquery/save-playlist.xqy";

    $.ajax({
          type: "GET",
          url: newurl, 
          data: { uri: PLAYLISTURI, plname : playlistName , gallery : galleryXml  },
          success: function(msg){
			try{
			     //Document is ready
			    $(function(){
                               // TODO only have to do this in case of new playlist
			       // add check 
                               populateLibraryListing(PLAYLISTURI+PLAYLISTDIR,"playlists", randomId());
                            });

			}catch(e)
			{
			    alert("ERROR"+e.description);
			}
		   }			
                   
     });
}

function deletePlaylist(playlistName)
{ 
    var newurl = DELETEPLAYLIST; //"/xquery/playlist-delete.xqy";
    $.ajax({
          type: "GET",
          url: newurl, 
          data: { uri: playlistName  },
          success: function(msg){
			try{
			     //Document is ready
			    $(function(){
                               populateLibraryListing(PLAYLISTURI+PLAYLISTDIR,"playlists", randomId());
                            });

			}catch(e)
			{
			    alert("ERROR"+e.description);
			}
		   }			
                   
     });
}


function publishPlaylist(playlistName)
{ 
    var newurl = EXPORTPLAYLIST; //"/xquery/export-playlist.xqy";

    $.ajax({
          type: "GET",
          url: newurl, 
          data: { plname : playlistName },
          success: function(msg){
			try{
			//need to poll for .pptx with unique id 
			//unique id provided by msg? or create here?
			//start poller here
		          intval="";	
			  start_int(msg);

			}catch(e)
			{
			    alert("ERROR"+e.description);
			}
		   }			
                   
     });
}
 
var modalWindow = {  
        parent:"body",  
        windowId:null,  
        content:null,  
        width:null,  
        height:null,  
        close:function(plname)  
        {    
            $(".modal-window").remove();  
            $(".modal-overlay").remove();

	    if(plname.length >0)
		  addPlaylist(plname);  
       },  
       open:function()  
       {  
          var modal = "";  
           modal += "<div class=\"modal-overlay\"></div>";  
           modal += "<div id=\"" + this.windowId + "\" class=\"modal-window\" style=\"width:" + this.width + "px; height:" + this.height + "px; margin-top:-" + (this.height / 2) + "px; margin-left:-" + (this.width / 2) + "px;\">";  
           modal += this.content;  
           modal += "</div>";      
     
          $(this.parent).append(modal);  
    
          $(".modal-window").append("<a class=\"close-window\"></a>");  

	  // add save button, close window, update for new playlist accordingly
           $(".close-window").click(function(){modalWindow.close();});  
           $(".modal-overlay").click(function(){modalWindow.close();});  
	   
       }  
};  

openMyModal = function(source)  
{  
    modalWindow.windowId = "myModal";  
    modalWindow.width = 480;  
    modalWindow.height = 405;  
    modalWindow.content = "<iframe width='480' height='205' frameborder='0' scrolling='no' allowtransparency='true' src='" + source + "'></iframe>";  
    modalWindow.open();  
};  

exportPlaylist = function(){
    var fullPlaylistName = $(".plname").text();
    var published = publishPlaylist(fullPlaylistName);
}

var intval="";
start_int = function(fileName){
     if(intval==""){
          intval=window.setInterval("start_poll('"+fileName+"')",2000);
     }else{
          stop_int(intval);
     }
}

stop_int = function(intval){
        
     if(intval!=""){
         window.clearInterval(intval);
         intval="";
     }
}


start_poll = function(fileName)
{ 
    var newurl = EXISTS; //"/xquery/exists.xqy";

    $.ajax({
          type: "GET",
          url: newurl, 
          data: { filename : fileName },
          success: function(msg){
			try{
			//need to poll for .pptx with unique id 
			//unique id provided by msg? or create here?

			   if(msg=="true"){
			      stop_int(intval);
			      openPptx(fileName);
			   }

			}catch(e)
			{
			    alert("ERROR"+e.description);
			}
		   }			
                   
     });
}

clearPlaylistExportLink = function()
{
    $('#openpptx').children().remove();
}

openPptx = function(fileName){

    //now question of whether to force export each time, or retain it if it exists in out?

    var tokens = fileName.split("/");
    var fName = tokens[tokens.length-1];

    var pptxLink = $('#openpptx');

    pptxLink.append("<a href='/slidetunes/xquery/open-binary.xqy?uri="+fileName+">"+
                                    fName+ 
                    "</a>");
	
/*
 * below is good and all, but we should provide a link and let browser decide what to do
	var newurl = OPENPPTX;
alert("FILENAME: "+fileName + "newurl: "+newurl);
	$.ajax({
          type: "GET",
          url: newurl, 
          data: { uri : fileName },
          success: function(msg){
			try{
			//need to poll for .pptx with unique id 
			//unique id provided by msg? or create here?
                            alert("DOC RETURNED");
			    return msg;

			    //DELETE NOW

			}catch(e)
			{
			    alert("ERROR"+e.description);
			}
		   }			
                   
     });
*/
}


