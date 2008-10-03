xquery version "1.0-ml";
declare namespace w    ="http://schemas.openxmlformats.org/wordprocessingml/2006/main";
declare namespace q    ="http://marklogic.com/beta/searchbox";
declare namespace xlink="http://www.w3.org/1999/xlink";
(
xdmp:set-response-content-type('text/html;charset=utf-8'),
'<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">',
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="content-type" content="text/html; charset=utf-8" />
	<title>Search and Reuse Sample</title>
	<link rel="stylesheet" type="text/css" href="../css/office-blue.css"/>
	<link rel="stylesheet" type="text/css" href="search.css"/>
	<script type="text/javascript" src="../js/prototype-1.6.0.3.js">//</script>
	<script type="text/javascript" src="../js/MarkLogicWordAddin.js">//</script>
	<script type="text/javascript" src="../js/debug.js">//</script>
	<script type="text/javascript" src="search.js">//</script>
</head>
<body>
	<div id="ML-Add-in">
	{
		let $q := xdmp:get-request-field("q") 
		let $intro := <div id="ML-Intro">
			<h1>Search and Reuse</h1>
			<p>Use the above search box to find content in Word 2007 documents stored on MarkLogic Server.
				Keywords narrow the results. Each search result represents a paragraph (or list item) that matches your criteria.</p>
			<p>To insert that entire paragraph into the active document at the current cursor location, double-click the result snippet.</p>
		</div>
		return (
			<form action="." method="get">
				<div class="ML-control">
					<div class="ML-label">
						<label for="ML-Search">Search</label>
					</div>
					<div class="ML-input">
						<input id="ML-Search" name="q" type="text" value="{$q}"/>
						<input id="ML-Submit" type="submit" value="Go"/>
					</div>
				</div>
			</form>,
			<ol id="ML-Results">{
				if($q) then
					let $tokens := tokenize($q, "\s+")
					let $and-query := cts:and-query(for $token in $tokens return cts:word-query($token, "case-insensitive"))
					let $or-query := cts:or-query(for $token in $tokens return cts:word-query($token, "case-insensitive"))
					let $hits := cts:search(//w:p, $and-query)
					let $start := (xs:int(xdmp:get-request-field("start")), 1)[1] 
					let $end := (xs:int(xdmp:get-request-field("end")), $start + 9)[1]
					return if(not($hits)) then
						(<div id="ML-Message"><p><strong>Nada, zilch, goose egg</strong></p><p>Your search for "{$q}" did not match anything.</p></div>,$intro)
					else
					for $hit in $hits[$start to $end]
						let $uri := base-uri($hit)
						let $snippet := if(string-length(data($hit)) > 120) then concat(substring(data($hit), 1, 250), "…") else data($hit)
						return <li title="{data($hit)}" xlink:href="{concat($uri,'#', 'xmlns(w=http://schemas.openxmlformats.org/wordprocessingml/2006/main) xpath(',xdmp:path($hit),')')}">
							{cts:highlight(<p>{$snippet}</p>, $or-query, <strong class="ML-highlight">{$cts:text}</strong>)}
							<ul class="ML-hit-metadata">
								<li>
									<a href="content.xqy?uri={xdmp:url-encode($uri)}" target="_blank">{substring-before(tokenize($uri,"/")[2],"_docx_parts")}</a>
								</li>
								<!-- <li>
									<a href="content.xqy?uri={xdmp:url-encode(replace($uri,'_docx_parts$','.docx'))}">{substring-before(tokenize($uri,"/")[2],"_docx_parts")}</a>
								</li> -->
							</ul>
						</li>
				else 
					$intro		
		}</ol>
		)
	}
		<div id="ML-Navigation">
			<a href="../">« Samples</a>
		</div>
	</div>
	</body>
</html>
)