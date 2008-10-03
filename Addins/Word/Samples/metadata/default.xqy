xquery version "0.9-ml"
declare namespace w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"

xdmp:set-response-content-type('text/html;charset=utf-8'),
'<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">',
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="content-type" content="text/html; charset=utf-8" />
	<title>Custom Metadata</title>
	<link rel="stylesheet" type="text/css" href="../css/office-blue.css"/>
	<link rel="stylesheet" type="text/css" href="metadata.css"/>
	<script type="text/javascript" src="../js/prototype-1.6.0.3.js">//</script>
	<script type="text/javascript" src="../js/MarkLogicWordAddin.js">//</script>
	<script type="text/javascript" src="../js/debug.js">//</script>
	<script type="text/javascript" src="metadata.js">//</script>
</head>
 <body>
	 <div id="ML-Add-in">
		<form id="ML-Metadata" action="." method="post">
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
		<!-- 
		<button id="ML-Save">Save</button>
		-->
		<button id="ML-Remove" class="ML-action" title="Remove the metadata from the current document">
			<img src="../img/delete.png"/> Remove
		</button>
		 
		<p id="ML-Message"></p>    
		<p id="v_fc2"></p>    
		                
		</div> 
		</form>
	  <div id="ML-Navigation">
			<a href="../">« Samples</a>
		</div>
	</div>
 </body>
</html>
