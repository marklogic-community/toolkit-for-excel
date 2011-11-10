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
