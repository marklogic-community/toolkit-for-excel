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
						(: Run the search unfiltered :)
						let $hits := cts:search(collection()[w:document], $and-query, "unfiltered")
						let $start := 1
						let $end := 10
						(: This doesn't actually truly paginate, becuase we're not filtering the hits :)
						let $results := for $hit in $hits (: [cts:contains(//w:p, $and-query)][1 to 10] :)
							(: Implement the filter :)
							let $paras := $hit//w:p[cts:contains(., $and-query)]
							return 
								if($paras) then
									<li><a href="asdf">{base-uri($hit)}</a>{
									for $para in $paras
										return cts:highlight(<p>{data($para)}</p>, $or-query, <strong>{$cts:text}</strong>)
									}</li>
								else ()
						return $results[1 to 2]
					else ()		
				}
				</ol>
				)
			}
		</body>
</html>
)