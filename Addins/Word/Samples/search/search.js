/* Copyright 2002-2008 Mark Logic Corporation.  All Rights Reserved. */

document.observe("dom:loaded", function() {
	// Listen to all double-click events on the ML-Results list container.
	// These will include double-clicks on its descendents.
	$("ML-Results").observe("dblclick", function(e) {
		// Get the parent list item
		var target = Event.findElement(e, "li");
		// Get the path to the document fragment from the result item's xlink:href attribute
		var path = target.readAttribute('xlink:href');
		// If there's no referece to external content, then we're done
		if(!path) return;
		
		// Request the actual WordprocessingML snippet from the server as an XLink+XPointer
		new Ajax.Request('content.xqy', 
			{
				'method': 'get',
				'parameters': {
					'uri': escape(path)
				},
				'requestHeaders': {
					'Accept': 'application/xml'
				},
				'onSuccess': function(response) {
					// Upon successful retrieval from the server,
					// insert the parsed XML into the active document
					MLA.insertBlockContent(response.responseXML);
				},
				'onFailure': function(response) {
					throw {
						message: "Failed AJAX response",
						payload: response
					}
				}
			}
		
		);
		// Clear the text selection that happens by default with a double-click
		document.selection.empty() ;
		// Don't bubble the event
		Event.stop(e);
	});
	
});