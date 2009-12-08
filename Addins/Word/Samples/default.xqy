xquery version "1.0-ml";
(:
Copyright 2008-2009 Mark Logic Corporation

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.

default.xqy - landing page for samples
:)

(
xdmp:set-response-content-type('text/html;charset=utf-8'),
'<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">',
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
			font-size: 200%;
			margin: 0.5em 0;
		}}
		p {{
			line-height: 1.45;
			margin: 0.5em 2em;
			font-size: 110%;
		}}
	</style>
	<script type="text/javascript" src="../js/prototype-1.6.0.3.js">//</script>
	<script type="text/javascript" src="../js/debug.js">//</script>
</head>
<body>
	<ul id="ML-Menu">
		<li>
			<h2><a href="search/">Search »</a></h2>
			<p>Search and explore Word 2007 content. Insert results into the active document.</p>
		</li>
		<li>
			<h2><a href="metadata/">Metadata »</a></h2>
			<p>Create and edit custom document metadata.</p>
		</li>
	</ul>
</body>
</html>
)
