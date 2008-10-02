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
	<title>Samples</title>
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
				<ol id="ML-Results">
				{
					if($q) then
						let $tokens := tokenize($q, "\s+")
						let $and-query := cts:and-query(for $token in $tokens return cts:word-query($token))
						let $or-query := cts:or-query(for $token in $tokens return cts:word-query($token))
						let $results := cts:search(collection()[w:document], $and-query, "unfiltered")
						let $start := 1
						let $end := 10
						for $result in $results[$start to $end]
							return <li><a href="content.xqy?uri={document-uri($result)}">asdf</a>
								<ol>{
									for $para in $result//w:p[cts:contains(., $or-query)]
									return cts:highlight(<li>{data($para)}</li>, $or-query, <strong>{$cts:text}</strong>)
								}</ol>
							</li>
					else ()		
				}
				</ol>
				)
			}
		</body>
</html>
)