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
	<script type="text/javascript" src="../js/prototype-1.6.0.3.js">//</script>
	<script type="text/javascript" src="../js/MarkLogicWordAddin.js">//</script>
	<script type="text/javascript" src="../js/debug.js">//</script>
	<script type="text/javascript" src="search.js">//</script>
</head>
<body>
			{
				let $q := xdmp:get-request-field("q") 
				return (
					<form action="." method="get">
						<div class="ML-control">
							<div class="ML-label">
								<label for="ML-Search">Search for</label>
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
							let $and-query := cts:and-query(for $token in $tokens return cts:word-query($token))
							let $or-query := cts:or-query(for $token in $tokens return cts:word-query($token))
							let $hits := cts:search(//w:p, $and-query)
							let $start := (xs:int(xdmp:get-request-field("start")), 1)[1] 
							let $end := (xs:int(xdmp:get-request-field("end")), $start + 9)[1]
							(: This doesn't actually truly paginate, becuase we're not filtering the hits :)
							for $hit in $hits[$start to $end]
								let $uri := base-uri($hit)
								return <li xlink:href="{concat($uri,'#', 'xmlns(w=http://schemas.openxmlformats.org/wordprocessingml/2006/main) xpath(',xdmp:path($hit),')')}">
									{cts:highlight(<p>{data($hit)}</p>, $or-query, <strong class="ML-highlight">{$cts:text}</strong>)}
									<ul class="ML-hit-metadata">
										<li>
											<a href="content.xqy?uri={xdmp:url-encode(concat($uri,'#xmlns(w=http://schemas.openxmlformats.org/wordprocessingml/2006/main) xpath(',xdmp:path($hit),')'))}" target="_blank">{substring-before(tokenize($uri,"/")[2],"_docx_parts")}</a>
										</li>
									</ul>
								</li>
						else ()		
					}</ol>
				)
			}
		</body>
</html>
)