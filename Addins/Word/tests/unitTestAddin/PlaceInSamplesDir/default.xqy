xquery version "0.9-ml"
declare namespace w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"

xdmp:set-response-content-type('text/html;charset=utf-8'),
'<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">',
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="content-type" content="text/html; charset=utf-8" />
	<title>Custom Metadata</title>
	<link rel="stylesheet" type="text/css" href="office-blue.css"/>
	<link rel="stylesheet" type="text/css" href="test.css"/>
	<script type="text/javascript" src="MarkLogicWordAddin.js">//</script>
	<script type="text/javascript" src="test.js">//</script>
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
		<button id="ML-Save" class="ML-action" onclick="updateMetadata(1)">Save</button>&nbsp;&nbsp;
		<button id="ML-Remove" class="ML-action" onclick="updateMetadata(2)" title="Remove the metadata from the current document">
			<img src="delete.png"/> Remove
		</button>
		 
		<p id="ML-Message"></p>  
		
		<div id="ML-Intro">
			<h1>Custom Metadata</h1>
			<p>Use the above form to manage custom metadata associated with the current document. 
			Upon saving the active document, this Dublin Core snippet is stored within the document’s <code>.docx</code> package as a XML document.</p>
			<p>Use the <code>Remove</code> button to discard the metadata.</p>
		</div>
		
				                
		</div> 
		</form>
	  <div id="ML-Navigation">
			<a href="default.xqy">« Samples</a>
		</div>
	</div>
 </body>
</html>
