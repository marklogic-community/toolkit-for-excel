xquery version "1.0-ml";

let $r := xdmp:set-response-content-type("text/html; charset=utf-8") 
let $doctype := '<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">'
return ( 
 $doctype,
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en" class="modal-html">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1"/> 
	<link href="../css/reset.css" type="text/css" rel="stylesheet" media="screen, print"/>	
	<link href="../css/jquery.hoverscroll.css" type="text/css" rel="stylesheet" media="screen" />
	<link href="../css/style.css" type="text/css" rel="stylesheet" media="screen, print"/>

	<title>modal</title>
</head>

<body class="modal-body">
<div class="modal-div">

<h1 class="modal-h1">Add Playlist</h1>
</div>
</body>
</html>
)
