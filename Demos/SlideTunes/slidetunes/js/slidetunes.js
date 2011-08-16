var RESTSERVER = "http://localhost:8060";
var APPSERVER = "http://localhost:8030";

var PRESODIR = "/paven/";
var PLAYLISTDIR = "/gallery/";

var SEARCHURI = RESTSERVER +"/office";
var PRESOURI = RESTSERVER+"/office/presentations";
var PLAYLISTURI = RESTSERVER+"/playlists";

var GETLIST = APPSERVER + "/shuffle2/get.xqy";
var GETSLDS = APPSERVER + "/shuffle2/get-slides.xqy";

$(document).ready(function() {

		$(document).bind("contextmenu",function(e){
                    return false;
                });
	 
		$('#deck-playlist ul').hoverscroll({
			width:		"100%",        // Width of the list
			height:		47         // Height of the list
		});
		
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
			opacity: 0.6
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

		populateLibraryListing(PRESOURI+PRESODIR, "presentations");
                populateLibraryListing(PLAYLISTURI+PLAYLISTDIR,"playlists");

		
	});

checkEventElement = function(e){
    var e=e ? e : window.event;
    var event_element=e.target? e.target : e.srcElement;
    return event_element;
}
	
populateLibraryListing = function(uri, destination){
    //alert(uri);
    var presos = simpleAjaxFetchPresentationList(uri, destination);
	
}

simpleAjaxFetchPresentationList = function(uri, destination)
{
    var newurl = GETLIST;

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
			//alert("Foo");
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


	    /*plList.append("<li>"+
                           "<div id='"+rId+"'>"+
 			       //"<a href='javascript:plAction("+'"'+pl+'"'+");'>"+
			            pl+
		               //"</a>"+
			    "</div>"+
		          "</li>");*/
         var aref = $('#'+rId);

	 aref.bind('mousedown', function(e) {
                        //setControlFocus(window.event.srcElement.id);
			//alert("Foo");
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
    //alert("rId in func: "+rId);
	
    var aref = $('#'+rId)

    var pos = aref.offset();  
    var width = aref.width();
        //show the menu directly over the placeholder
    $(".vmenu").css( { "left": (pos.left) + "px", "top":pos.top + "px" } );
    $(".vmenu").show();
 
        //need to unbind, or we keep binding and end up with multiple 
    $('.vmenu .first_li').bind('click',function() {
        //plAction()
	plAction(aref.text());
        $('.vmenu .first_li').unbind('click');
        $('.vmenu').hide();
    });

    $('.vmenu .second_li').bind('click',function() {
        //alert("SECOND"+aref.text() + $('.vmenu .second_li').text() );
	presoAction(aref.text());
        $('.vmenu .second_li').unbind('click');
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
}

presoAction = function(presentation){
    var serveruri = PRESOURI;
    var slideuri = presentation + "/slides";
    alert("PRESOACTION: "+slideuri);
    //simpleAjaxFetchImages(serveruri, slideuri, "workspace");
}

plAction = function(playlist){
    var serveruri = PRESOURI;
    var slideuri = PLAYLISTURI+playlist;
    alert("PLACTION: "+slideuri);
    //simpleAjaxFetchImages(serveruri, slideuri, "playlists");
}


