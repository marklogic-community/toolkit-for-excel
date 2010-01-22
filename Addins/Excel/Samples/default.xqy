xquery version "1.0-ml";
(:
Copyright 2008-2010 Mark Logic Corporation

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
:)

(
xdmp:set-response-content-type('text/html;charset=utf-8'),
'<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">',
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="content-type" content="text/html; charset=utf-8" />
	<title>Samples</title>
	<link rel="stylesheet" type="text/css" href="css/office-blue.css"/>
	<style type="text/css">
		body {{
			padding: 2em 4em;
		}}
		h2 {{
			font-size: 150%;
			margin: 0.5em 0;
		}}
		p {{
			line-height: 1.45;
			margin: 0.5em 2em;
			font-size: 90%;
		}}
	</style>
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
