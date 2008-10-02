/* $Id: search.js,v 1.2 2008-10-02 22:37:48 jmakeig Exp $ */

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

var MarkLogic = window.MarkLogic || {}

MLA.OOXML = {}
MLA.OOXML.namespaces = {
		"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
}
MLA.OOXML.createParagraph = function(text) {
	return "<w:p xmlns:w='"+MLA.OOXML.namespaces.w+"'><w:r><w:t>"+text+"</w:t></w:r></w:p>";
}

// EXPERIMENTAL
MarkLogic.DOM = function() {
	var MSXML_VERSION = 'Msxml2.DOMDocument.6.0';
	return {
		parse: function(xml, options) {
			var dom = new ActiveXObject(MSXML_VERSION);
			dom.async = false;
			dom.validateOnParse = false;
			dom.resolveExternals = false;
			options = options || {}
			for(p in options) { dom[p] = options[p] }
			if(xml) dom.loadXML(xml);
			return dom;
		}
	}
}();

MLA.DOM = {}
MLA.DOM.insertBlockContent = function(node) {
	MLA.insertBlockContent(node.xml);
}