/* $Id: search.js,v 1.3 2008-10-02 22:38:56 jmakeig Exp $ */

document.observe("dom:loaded", function() {
	$("ML-Results").observe("dblclick", function(e) {
		var target = Event.findElement(e, "li");
		// This is wrong in so many ways. Should be using getAttributeNS
		var path = target.readAttribute('xlink:href');
		_l(path);
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
					MLA.insertBlockContent(response.responseXML);
				},
				'onFailure': function(response) {
					throw {
						message: "Failed AJAX response",
						payload: response
					}
				},
				'onComplete': function() {
				}
			}
		
		);
		
		//COMPAT: Selection/range handling varies across browsers
		document.selection.empty() ;
		//window.getSelection().removeAllRanges();
		Event.stop(e);
	});
	
});