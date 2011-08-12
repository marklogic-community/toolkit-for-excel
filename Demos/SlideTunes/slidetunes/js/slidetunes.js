	$(document).ready(function() {
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
		
	});

updateLibPresentationList = function(msg){

    try{
        var local =  MLA.createXMLDOM(msg);
        var presos = local.getElementsByTagName("presentation");

	if($('#presolist').children('li').length){   
	    $('#presolist').children('li').remove();
     	}

	var presoList = $('#presolist');

        for( var i = 0; i < presos.length; i++) {
            var pres =  presos[i].childNodes[0].nodeValue;
	    var rId = "pid"+i;
	    presoList.append("<li>"+
		               "<div id='"+rId+"'>"+
			        //  "<a href='javascript:presoAction("+'"'+pres+'"'+");'>"+
			              pres+
				//   "</a>"+
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

