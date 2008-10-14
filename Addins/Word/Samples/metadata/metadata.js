document.observe(
	"dom:loaded",
	function() {
		/**
		 * Generate the Dublin Core document.
		 * @param {String} title
		 * @param {String} description
		 * @param {String} publisher
		 * @param {String} id
		 * @return {String} The text serialization of the XML
		 */
		function generateTemplate(title, description, publisher, id) {
			var v_template = "<metadata "
					+ "xmlns='http://example.org/myapp/' "
					+ "xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' "
					+ "xsi:schemaLocation='http://example.org/myapp/ http://example.org/myapp/schema.xsd' "
					+ "xmlns:dc='http://purl.org/dc/elements/1.1/'>"
					+ "<dc:title>" + title + "</dc:title>"
					+ "<dc:description>" + description
					+ "</dc:description>" + "<dc:publisher>"
					+ publisher + "</dc:publisher>"
					+ "<dc:identifier>" + id + "</dc:identifier>"
					+ "</metadata>";
			return v_template;
	
		}
		/**
		 * Save the metadata from the form as an XML document in the active document.
		 */
		function updateMetadata() {
			showMessage("Saving metadata…", -1);
			var edited = false;
			_l("Saving Custom Piece");
			try{
			var customPieceIds = MLA.getCustomPieceIds();
			}catch(ex){
				// TODO: Figure out why this exception is being thrown when Word closes. Probably some timing thing about the order of things eing unloaded
				return;
			}
			_l(customPieceIds.length);
			var customPieceId = null;
			for (i = 0; i < customPieceIds.length; i++) {
				if (customPieceIds[i] == null
						|| customPieceIds == "") {
					// do nothing
				} else {
					customPieceId = customPieceIds[i];
					var delPiece = MLA
							.deleteCustomPiece(customPieceId);
					_l("Deleted " + delPiece);
					edited = true;
				}
			}
			
			// Escape hatch for when the DOM has been unloaded 
			if(!$("ML-Title")) 
				return;
			
			_l("getting values");
			var v_title = $("ML-Title").value;
			_l(v_title);
			var v_description = $("ML-Desc").value;
			_l(v_description);
			var v_publisher = $("ML-Publisher").value;
			_l(v_publisher);
			var v_identifier = $("ML-Id").value;
			_l(v_identifier);
			
			var customPiece = generateTemplate(v_title,
					v_description, v_publisher, v_identifier);
	
			_l(customPiece);
	
			var newid = MLA.addCustomPiece(customPiece);
	
			if (edited) {
				_l("Metadata Edited");
			}
			showMessage("Metadata saved");
		}
		
		// Default the status message to hidden
		$("ML-Message").setStyle({display: "none"});
		
		/**
		 * Show a GMail-style status message.
		 * 
		 * @param {String} message The text message
		 * @param {Number} The duration that the message should be visible. Defaults to 1000ms. < 0 means infinite.
		 */
		function showMessage(message, duration) {
			duration = duration || 1000;
			$("ML-Message").setStyle({display: "block"});
			$("ML-Message").innerHTML = message;
			
			if(duration > 0) {
				setTimeout( function() {
					$("ML-Message").innerHTML = "";
					$("ML-Message").setStyle({display: "none"});
				}, duration);
			}
		}
		function removeMetadata() {
			_l("Removing Custom Piece");
			var customPieceIds = MLA.getCustomPieceIds();
			var customPieceId = null;
			for (i = 0; i < customPieceIds.length; i++) {
				if (customPieceIds[i] == null
						|| customPieceIds == "") {
					// do nothing
				} else {
					customPieceId = customPieceIds[i];
					var delPiece = MLA
							.deleteCustomPiece(customPieceId);
				}
	
			}
			[ "ML-Title", "ML-Desc", "ML-Publisher", "ML-Id" ].each(function(el) {
				$(el).value = "";
			});
			showMessage("Metadata removed");
		}
		
		$("ML-Remove").observe("click",function(e) {
			removeMetadata()
		});
		
		// Cancel the default form submission. This should never be called becuase there's no submit button in the UI.
		$("ML-Metadata").observe("submit", function(e) {
			Event.stop(e);
			return false;
		});
		
		// Save after losing focus on each control
		// 'change' would be a better event, but it doesn't get fired when focus leaves the add-in
		[ "ML-Title", "ML-Desc", "ML-Publisher", "ML-Id" ].each(function(el) {
			// FIXME: Both 'blur' and 'change' disable the tab key's ability to switch between fields
			$(el).observe("blur", function(e) {
				updateMetadata();
			});
		});
		
		
		var customPieceIds =  MLA.getCustomPieceIds();
		var customPieceId = null;
		var tmpCustomPieceXml = null;
		
		for (i = 0; i < customPieceIds.length; i++) {
			if (customPieceIds[i] == null || customPieceIds == "") {
				// do nothing
			} else {
				_l("PIECE ID: " + customPieceIds[i]);
				customPieceId = customPieceIds[i];
				tmpCustomPieceXml = MLA.getCustomPiece(customPieceId);
				_l(tmpCustomPieceXml.xml);
			}
		}
		
		var xmlDoc = tmpCustomPieceXml;
		if(xmlDoc) {
			_l(xmlDoc.xml);
			var v_title="";
			var v_description="";
			var v_publisher="";
			var v_identifier="";
			if(xmlDoc.getElementsByTagName("dc:title")[0].hasChildNodes()) 
			   v_title = xmlDoc.getElementsByTagName("dc:title")[0].childNodes[0].nodeValue;

			if(xmlDoc.getElementsByTagName("dc:description")[0].hasChildNodes()) 
			   v_description = xmlDoc.getElementsByTagName("dc:description")[0].childNodes[0].nodeValue;

			if(xmlDoc.getElementsByTagName("dc:publisher")[0].hasChildNodes()) 
			   v_publisher = xmlDoc.getElementsByTagName("dc:publisher")[0].childNodes[0].nodeValue;
			
			if(xmlDoc.getElementsByTagName("dc:identifier")[0].hasChildNodes()) 
			   v_identifier = xmlDoc.getElementsByTagName("dc:identifier")[0].childNodes[0].nodeValue;

			$("ML-Title").value = v_title;
			$("ML-Desc").value = v_description;
			$("ML-Publisher").value = v_publisher;
			$("ML-Id").value = v_identifier;
		}
	}
);
