xquery version "1.0-ml";
(
xdmp:set-response-content-type('text/html;charset=utf-8'),
'<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">',
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="content-type" content="text/html; charset=utf-8" />
	<title>Test JS</title>
	<link rel="stylesheet" type="text/css" href="css/office-blue.css"/>
	<script type="text/javascript" src="js/MarkLogicExcelAddin.js">//</script>
	<script type="text/javascript" src="js/test.js">//</script>
</head>
<body>
	<ul id="ML-Menu">
                <li>
			<!--<h2><a href="tabletoexcel.xqy">Search »</a></h2>-->			<h2><a href="search/default.xqy">Search »</a></h2>
			<p>Search and explore Excel 2007 and other XML content.  Open existing Workbooks or Create a New Excel Worksheet from an XML Table.</p>
		</li>
		<li>
			<h2><a href="metadata/metadata.xqy">Metadata »</a></h2>
			<p>Create and edit custom document metadata.</p>
		</li>
                <li>
			<!-- <h2><a href="openexcel.xqy">Save »</a></h2> -->
			<h2><a href="save/save.xqy">Save »</a></h2>
			<p>Save the Active Workbook directly to MarkLogic.</p>
		</li>

	</ul>
</body>
</html>
)
