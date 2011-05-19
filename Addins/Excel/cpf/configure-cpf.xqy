xquery version "1.0-ml";

declare namespace config="http://marklogic.com/toolkit/excel/cpf-config";
import module namespace dom = "http://marklogic.com/cpf/domains" 
		  at "/MarkLogic/cpf/domains.xqy";
import module namespace p = "http://marklogic.com/cpf/pipelines" at "/MarkLogic/cpf/pipelines.xqy";

declare variable $config:Support-Src-Path := 
          "C:\Users\paven\Desktop\installer\excel\xquery\";
declare variable $config:Support-Dest-Path := 
          "C:\Program Files\MarkLogic\Modules\MarkLogic\openxml\";

declare variable $config:CPF-Source-Path :=
          "C:\Users\paven\Desktop\installer\excel\cpf\";
declare variable $config:CPF-Dest-Path :=
          "C:\Program Files\MarkLogic\Modules\MarkLogic\conversion\actions\";


declare variable $config:Triggers-Database := "TK-Triggers";

declare variable $config:Domain-Name := "OpenXML";
declare variable $config:Domain-Description := "Handling incoming Open XML documents";
declare variable $config:Domain-Scope := "directory";
declare variable $config:Domain-URI := "/";
declare variable $config:Domain-Depth := "infinity";
declare variable $config:Context-Database := "Modules";
declare variable $config:Context-Root :=     "/";
declare variable $config:Restart-User := "admin";


declare function config:install-xqy-support()
{
try{
  let $support-file := "spreadsheet-ml-support.xqy"
  let $cpf-file := "map-shared-action.xqy"
  let $src-support  := fn:concat($config:Support-Src-Path, $support-file)
  let $dest-support := fn:concat($config:Support-Dest-Path, $support-file)
  let $support-doc  := xdmp:document-get($src-support)
  let $src-cpf := fn:concat($config:CPF-Source-Path,$cpf-file)
  let $dest-cpf:= fn:concat($config:CPF-Dest-Path,$cpf-file)
  let $cpf-doc := xdmp:document-get($src-cpf)

  return (xdmp:save($dest-support, $support-doc),
          xdmp:save($dest-cpf, $cpf-doc),
          fn:concat("Step 1: install-xqy-support(); ",$support-file," copied to ",$config:Support-Dest-Path),
          fn:concat("Step 1: install-xqy-support(); ",$cpf-file," copied to ",$config:CPF-Dest-Path)
          )

}catch($e){
  $e
}
};


declare function config:create-domain(){

try{
xdmp:eval("
           declare namespace config='http://marklogic.com/toolkit/excel/cpf-config';
           import module namespace dom = 'http://marklogic.com/cpf/domains' 
		  at '/MarkLogic/cpf/domains.xqy';
           import module namespace p = 'http://marklogic.com/cpf/pipelines' at '/MarkLogic/cpf/pipelines.xqy';
           declare variable $config:d-name as xs:string external;
           declare variable $config:triggersdb as xs:string external;
           declare variable $config:d-description as xs:string external;
           declare variable $config:d-scope as xs:string external;
           declare variable $config:d-URI as xs:string external;
           declare variable $config:d-depth as xs:string external;
           declare variable $config:c-db as xs:string external;
           declare variable $config:c-root as xs:string external;
              dom:create( $config:d-name, $config:d-description, 
              dom:domain-scope( $config:d-scope, 
                    $config:d-URI, 
		    $config:d-depth),
	      dom:evaluation-context( xdmp:database($config:c-db), 
		                     $config:c-root ),
              (), 
              () )", 
              ((xs:QName("config:d-name"), $config:Domain-Name),
               (xs:QName("config:d-description"), $config:Domain-Description),
               (xs:QName("config:d-scope"), $config:Domain-Scope),
               (xs:QName("config:d-URI"), $config:Domain-URI),
               (xs:QName("config:d-depth"), $config:Domain-Depth),
               (xs:QName("config:d-description"), $config:Domain-Description),
               (xs:QName("config:c-db"), $config:Context-Database),
               (xs:QName("config:c-root"), $config:Context-Root)
               ),
              <options xmlns="xdmp:eval">
		    <database>{xdmp:database($config:Triggers-Database)}</database>
	      </options>),
              "Step 2: create-domain(); Domain Created"
}catch($e){
 if(fn:string($e/error:code) eq "CPF-DOMAINEXISTS")then
   "Step 2: create-domain() encountered the following: CPF-DOMAINEXISTS - Domain Already Defined"
 else fn:string($e/error:code)
}

};

declare function config:create-configuration()
{
try{
xdmp:eval("declare namespace config='http://marklogic.com/toolkit/excel/cpf-config';
           import module namespace dom = 'http://marklogic.com/cpf/domains' 
		  at '/MarkLogic/cpf/domains.xqy';
           import module namespace p = 'http://marklogic.com/cpf/pipelines' at '/MarkLogic/cpf/pipelines.xqy';
           declare variable $config:r-user as xs:string external;
           declare variable $config:c-db as xs:string external;
           declare variable $config:d-name as xs:string external;

              dom:configuration-create( $config:r-user, 
              dom:evaluation-context( xdmp:database($config:c-db), '/' ),
              fn:data(dom:get($config:d-name)/dom:domain-id), 
              ())",
              ((xs:QName("config:r-user"),$config:Restart-User),
               (xs:QName("config:c-db"), $config:Context-Database),
               (xs:QName("config:d-name"), $config:Domain-Name)
              ),
              <options xmlns="xdmp:eval">
		    <database>{xdmp:database($config:Triggers-Database)}</database>
              </options>),
         "Step 3: create-configuration(); Configuration Created\n"
}catch($e){
 if(fn:string($e/error:code) eq "CPF-CONFIGEXISTS")then
   "Step 3: create-configuration() encountered the following: CPF-CONFIGEXISTS - Configuration Already Exists"
 else fn:string($e/error:code)
}


};

declare function config:load-pipelines()
{
try{
xdmp:eval("declare namespace config='http://marklogic.com/toolkit/excel/cpf-config';
           import module namespace dom = 'http://marklogic.com/cpf/domains' 
		  at '/MarkLogic/cpf/domains.xqy';
           import module namespace p = 'http://marklogic.com/cpf/pipelines' at '/MarkLogic/cpf/pipelines.xqy';
           declare variable $config:cpf-path as xs:string external;
           
              (p:insert(xdmp:document-get('Installer/cpf/status-pipeline.xml')),
               p:insert(xdmp:document-get('Installer/openxml/openxml-pipeline.xml')),
               p:insert(xdmp:document-get(fn:concat($config:cpf-path,'/spreadsheetml-pipeline.xml'))))",
               ((xs:QName("config:cpf-path"), $config:CPF-Source-Path)),
               <options xmlns="xdmp:eval">
		    <database>{xdmp:database($config:Triggers-Database)}</database>
               </options>),
 "Step 4: load-pipelines(); Status Change Handling, Open XML Extract, and SpreadsheetML Process pipelines loaded"
}catch($e){
 $e
}

};

declare function config:attach-pipelines()
{
try{
xdmp:eval("declare namespace config='http://marklogic.com/toolkit/excel/cpf-config';
           import module namespace dom = 'http://marklogic.com/cpf/domains' 
		  at '/MarkLogic/cpf/domains.xqy';
           import module namespace p = 'http://marklogic.com/cpf/pipelines' at '/MarkLogic/cpf/pipelines.xqy';
           declare variable $config:d-name as xs:string external;
              (dom:add-pipeline( $config:d-name, 
		  p:get('Status Change Handling')/p:pipeline-id ),
               dom:add-pipeline( $config:d-name, 
		  p:get('Office OpenXML Extract')/p:pipeline-id ),
               dom:add-pipeline( $config:d-name, 
		  p:get('SpreadsheetML Process')/p:pipeline-id ))",
               ((xs:QName("config:d-name"), $config:Domain-Name)),
               <options xmlns="xdmp:eval">
		    <database>{xdmp:database($config:Triggers-Database)}</database>
               </options>),
 "Step 5: attach-pipelines(); Pipelines attached to domain"
}catch($e){
$e
}
};
config:install-xqy-support(),
config:create-domain(),
config:create-configuration(),
config:load-pipelines(),
config:attach-pipelines()

