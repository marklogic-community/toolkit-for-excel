window.onload = initPage;

function initPage() {

	_l("initializing page");

	var customPieceIds = MLA.getCustomPieceIds();
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

	if (tmpCustomPieceXml != null)// && tmpCustomPieceXml.length > 1)
	{
		_l("IN IF");
		var xmlDoc = tmpCustomPieceXml;
		// xmlDoc=new ActiveXObject("Microsoft.XMLDOM");
		// xmlDoc.async="false";
		// xmlDoc.loadXML(tmpCustomPieceXml);
		var v_title = xmlDoc.getElementsByTagName("dc:title")[0].childNodes[0].nodeValue;
		var v_description = xmlDoc.getElementsByTagName("dc:description")[0].childNodes[0].nodeValue;
		var v_publisher = xmlDoc.getElementsByTagName("dc:publisher")[0].childNodes[0].nodeValue;
		var v_identifier = xmlDoc.getElementsByTagName("dc:identifier")[0].childNodes[0].nodeValue;
		$("ML-Title").value = v_title;
		$("ML-Desc").value = v_description;
		$("ML-Publisher").value = v_publisher;
		$("ML-Id").value = v_identifier;
		document.getElementById("v_fc").innerHTML = "Metadata Saved with Document";

	} else {
		document.getElementById("v_fc").innerHTML = "No Metadata Saved with Document";
	}

}

function generateTemplate(title, description, publisher, id) {
	var v_template = "<metadata "
			+ "xmlns='http://example.org/myapp/' "
			+ "xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' "
			+ "xsi:schemaLocation='http://example.org/myapp/ http://example.org/myapp/schema.xsd' "
			+ "xmlns:dc='http://purl.org/dc/elements/1.1/'>" + "<dc:title>"
			+ title + "</dc:title>" + "<dc:description>" + description
			+ "</dc:description>" + "<dc:publisher>" + publisher
			+ "</dc:publisher>" + "<dc:identifier>" + id + "</dc:identifier>"
			+ "</metadata>";
	return v_template;

}

function updateMetadata(i) {
	var edited = false;
	if (i == 1) {

		_l("Saving Custom Piece");

		var customPieceIds = MLA.getCustomPieceIds();
		_l(customPieceIds.length);
		var customPieceId = null;
		for (i = 0; i < customPieceIds.length; i++) {
			if (customPieceIds[i] == null || customPieceIds == "") {
				// do nothing
			} else {
				customPieceId = customPieceIds[i];
				var delPiece = MLA.deleteCustomPiece(customPieceId);
				_l("Deleted " + delPiece);
				edited = true;
			}

		}
		_l("getting values");
		var v_title = $("ML-Title").value;
		_l(v_title);
		var v_description = $("ML-Desc").value;
		_l(v_description);
		var v_publisher = $("ML-Publisher").value;
		_l(v_publisher);
		var v_identifier = $("ML-Id").value;
		_l(v_identifier);
		/*
		 * if(v_title=="" || v_title==null) v_title="Please Enter A Title";
		 * if(v_description=="" || v_description==null) v_description="Please
		 * Enter A Description"; if(v_publisher=="" || v_publisher==null)
		 * v_publisher="Please Enter A Publisher"; if(v_identifier=="" ||
		 * v_identifier==null) v_identifier="Please Enter An Id";
		 */
		var customPiece = generateTemplate(v_title, v_description, v_publisher,
				v_identifier);

		_l(customPiece);

		var newid = MLA.addCustomPiece(customPiece);

		if (edited) {
			_l("Metadata Edited");
		}
	} else {
		_l("Removing Custom Piece");
		var customPieceIds = MLA.getCustomPieceIds();
		var customPieceId = null;
		for (i = 0; i < customPieceIds.length; i++) {
			if (customPieceIds[i] == null || customPieceIds == "") {
				// do nothing
			} else {
				customPieceId = customPieceIds[i];
				var delPiece = MLA.deleteCustomPiece(customPieceId);
			}

		}

	}
}
