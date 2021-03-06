xquery version "1.0-ml";
(:
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
:)

xdmp:set-response-content-type('text/html;charset=utf-8'),
'<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">',
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="content-type" content="text/html; charset=utf-8" />
	<title>Custom Metadata</title>
	<link rel="stylesheet" type="text/css" href="../css/office-blue.css"/>
	<link rel="stylesheet" type="text/css" href="metadata.css"/>
	<script type="text/javascript" src="../js/MarkLogicExcelAddin.js">//</script>
	<script type="text/javascript" src="metadata.js">//</script>
</head>
 <body>
	 <div id="ML-Add-in">
		<form id="ML-Metadata" action="." method="get">
			<div class="ML-control">
				<div class="ML-label">
					<label for="ML-Title">Title</label>
				</div>
				<div class="ML-input">
					<input id="ML-Title"/>
				</div>
			</div>
			<div class="ML-control">
				<div class="ML-label">
					<label for="ML-Desc">Description</label>
				</div>
				<div class="ML-input">
					<textarea id="ML-Desc"></textarea>
				</div>
			</div>
			<div class="ML-control">
				<div class="ML-label">
					<label for="ML-Publisher">Publisher</label>
				</div>
				<div class="ML-input">
					<input id="ML-Publisher"/>
				</div>
			</div>
			<div class="ML-control">
				<div class="ML-label">
					<label for="ML-Id">Identifier</label>
				</div>
				<div class="ML-input">
					<input id="ML-Id"/>
				</div>
			</div>
		<div>
		<button id="ML-Save" class="ML-action" onclick="updateMetadata(1)">Add Metadata</button>&nbsp;&nbsp;
		<button id="ML-Remove" class="ML-action" onclick="updateMetadata(2)" title="Remove the metadata from the current document">
			<img src="delete.png"/> Remove
		</button>
		 
	
		
		<div id="ML-Intro">
			<h1>Custom Metadata</h1>
			<p>Use the above form to manage custom metadata associated with the current document. 
			Upon saving the active document, this Dublin Core snippet is stored within the document’s <code>.xlsx</code> package as a XML document.</p>
			<p>Use the <code>Remove</code> button to discard the metadata.</p>
		</div>

                <br/>
                <br/>
                <p id="ML-Message"></p> 
		
				                
		</div> 
		</form>
	  <div id="ML-Navigation">
			<a href="../default.xqy">« Samples</a>
		</div>
	</div>
 </body>
</html>
