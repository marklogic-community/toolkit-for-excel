xquery version "1.0-ml";

module namespace requests="http://marklogic.com/appservices/requests";

import module namespace rest = "http://marklogic.com/appservices/rest"
        at "/MarkLogic/appservices/utils/rest.xqy";

declare variable $requests:options as element(rest:options)
    :=
<rest:options>
  <rest:request uri="^(.+)/slide(\d+)$" endpoint="/slide-image.xqy">
    <rest:uri-param name="deck">$1</rest:uri-param>
    <rest:uri-param name="slide" as="integer">$2</rest:uri-param>
    <rest:http method="GET"/>    
    <rest:param name="size" match="(small|medium|large)">$1</rest:param>
    (:<rest:param name="format" match="(json|xml)">$1</rest:param>:)
  </rest:request>
  <rest:request uri="^(.+)/presentations(.+/)$" endpoint="/presentations.xqy">
    <rest:uri-param name="directory">$2</rest:uri-param>
    <rest:http method="GET"/>
    <rest:param name="format" match="(xml|json)">$1</rest:param>
    <rest:param name="start" as="integer">$1</rest:param>
  </rest:request>
  <rest:request uri="^(.+)/presentations(.+/)$" endpoint="/directory-create.xqy">
    <rest:uri-param name="directory">$2</rest:uri-param>
    <rest:http method="PUT"/>
  </rest:request>
  <rest:request uri="^(.+)/presentations(.+/)$" endpoint="/directory-delete.xqy">
    <rest:uri-param name="directory">$2</rest:uri-param>
    <rest:http method="DELETE"/>
  </rest:request>
  <rest:request uri="^(.+)/slides$" endpoint="/slide-uris.xqy">
    <rest:uri-param name="deck">$1</rest:uri-param>
    <rest:http method="GET"/>
    <rest:param name="format" match="(xml|json)">$1</rest:param>
  </rest:request>
  <rest:request uri="^(.+presentations.+(ppt|pptx))$" endpoint="/presentation-fetch.xqy">
    <rest:uri-param name="deck">$1</rest:uri-param>
    <rest:http method="GET"/>
  </rest:request>
  <rest:request uri="^(.+(office))$" endpoint="/search.xqy">
    <rest:uri-param name="deck">$1</rest:uri-param>
    <rest:http method="GET"/>    
    <rest:param name="q">$1</rest:param>
    <rest:param name="format" match="(xml|json)">$1</rest:param>
    <rest:param name="start" as="integer">$1</rest:param>
  </rest:request>
  <rest:request uri="^(.+)playlists(.+/)$" endpoint="/playlists.xqy">
    <rest:uri-param name="directory">$2</rest:uri-param>
    <rest:http method="GET"/>
    <rest:param name="format" match="(xml|json)">$1</rest:param>
    <rest:param name="start" as="integer">$1</rest:param>
  </rest:request>
  <rest:request uri="^(.+playlists.+(xml|json))$" endpoint="/playlist-fetch.xqy">
    <rest:uri-param name="deck">$1</rest:uri-param>
    <rest:http method="GET"/>
  </rest:request>
  <rest:request uri="^(.+playlists.+(xml|json))$" endpoint="/playlist-put.xqy">
    <rest:uri-param name="deck">$1</rest:uri-param>
    <rest:http method="PUT"/>
  </rest:request>
 <rest:request uri="^(.+playlists.+(xml|json))$" endpoint="/playlist-delete.xqy">
    <rest:uri-param name="deck">$1</rest:uri-param>
    <rest:http method="DELETE"/>
  </rest:request>
</rest:options>;
