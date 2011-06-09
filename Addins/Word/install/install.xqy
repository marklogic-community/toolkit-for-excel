xquery version "1.0-ml";


declare namespace config="http://marklogic.com/toolkit/word/cpf-config";
import module namespace dom = "http://marklogic.com/cpf/domains" 
		  at "/MarkLogic/cpf/domains.xqy";
import module namespace p = "http://marklogic.com/cpf/pipelines" at "/MarkLogic/cpf/pipelines.xqy";

(: Source and target directories for word-processing-ml-support.xqy :)
declare variable $config:SUPPORT-SRC-PATH := 
          "C:\Users\paven\Desktop\Word\xquery\";

(: Support .xqy files will be copied from $config:Support-Src-Path
   and placed under $config:Server-Root/Modules/MarkLogic/openxml/".
   If you have a non-standard configuration of MarkLogic Server,  
   update $config:Server-Root accordingly. :)
declare variable $config:SERVER-ROOT:= 

                  let $platform := xdmp:platform()
                  return if($platform eq "winnt") then
                               "C:/Program Files/MarkLogic/"
                         else if($platform eq "linux") then
                               "/opt/MarkLogic/"
                         else if($platform eq "macosx") then
                               "~/Library/MarkLogic/"
                         else if($platform eq "solaris") then
                               "/opt/MARKlogic/"
                         else ();

(: The CPF Restart User. The default below assumes a user named 'admin' 
   having admin priveleges.  Update for your environment accordingly.

   Also, this is fine for development
   but you'll want to reconsider your restart user when deploying your 
   application to a Production environment.:)   
declare variable $config:RESTART-USER := "admin";

declare variable $config:TRIGGERS-DB := "TK-Triggers";

(:  If you already have CPF installed with a Domain configured for the same $config:Domain-URI
    you'll end up with 2 domains with different names, but for the same URI. 
    Don't cross the streams!  Rename $Domain-Name accordingly. :)
declare variable $config:DOMAIN-NAME := "OpenXML";
declare variable $config:DOMAIN-DESCRIPTION := "Handling incoming Open XML documents";
declare variable $config:DOMAIN-SCOPE := "directory";
declare variable $config:DOMAIN-URI := "/";
declare variable $config:DOMAIN-DEPTH := "infinity";
declare variable $config:CONTEXT-DB := "Modules";
declare variable $config:CONTEXT-ROOT :=     "/";


declare function config:install-xqy-support()
{
try{
  let $file := "word-processing-ml-support.xqy"
  let $src  := fn:concat($config:SUPPORT-SRC-PATH, $file)
  let $openxml-path := fn:concat($config:SERVER-ROOT,"Modules/MarkLogic/openxml/")
  let $dest := fn:concat($openxml-path, $file)
  let $doc  := xdmp:document-get($src)
  return (xdmp:save($dest, $doc),
         fn:concat("Step 1: install-xqy-support(); ",$file," copied to ",$openxml-path))
         
}catch($e){
  $e
}
};


declare function config:create-domain(){

try{
xdmp:eval("
           declare namespace config='http://marklogic.com/toolkit/word/cpf-config';
           import module namespace dom = 'http://marklogic.com/cpf/domains' 
		  at '/MarkLogic/cpf/domains.xqy';
           import module namespace p = 'http://marklogic.com/cpf/pipelines' at '/MarkLogic/cpf/pipelines.xqy';
           declare variable $config:d-name as xs:string external;
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
              ( (xs:QName("config:d-name"), $config:DOMAIN-NAME),
                (xs:QName("config:d-description"), $config:DOMAIN-DESCRIPTION),
                (xs:QName("config:d-scope"), $config:DOMAIN-SCOPE),
                (xs:QName("config:d-URI"), $config:DOMAIN-URI),
                (xs:QName("config:d-depth"), $config:DOMAIN-DEPTH),
                (xs:QName("config:c-db"), $config:CONTEXT-DB),
                (xs:QName("config:c-root"), $config:CONTEXT-ROOT)
               ),
              <options xmlns="xdmp:eval">
		    <database>{xdmp:database($config:TRIGGERS-DB)}</database>
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
xdmp:eval("declare namespace config='http://marklogic.com/toolkit/word/cpf-config';
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
              ((xs:QName("config:r-user"),$config:RESTART-USER),
               (xs:QName("config:c-db"), $config:CONTEXT-DB),
               (xs:QName("config:d-name"), $config:DOMAIN-NAME)
               
               ),
              <options xmlns="xdmp:eval">
		    <database>{xdmp:database($config:TRIGGERS-DB)}</database>
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
xdmp:eval("declare namespace config='http://marklogic.com/toolkit/word/cpf-config';
           import module namespace dom = 'http://marklogic.com/cpf/domains' 
		  at '/MarkLogic/cpf/domains.xqy';
           import module namespace p = 'http://marklogic.com/cpf/pipelines' at '/MarkLogic/cpf/pipelines.xqy';
           
              (p:insert(xdmp:document-get('Installer/cpf/status-pipeline.xml')),
               p:insert(xdmp:document-get('Installer/openxml/openxml-pipeline.xml')),
               p:insert(xdmp:document-get('Installer/openxml/wordprocessingml-pipeline.xml')))",
               (),
               <options xmlns="xdmp:eval">
		    <database>{xdmp:database($config:TRIGGERS-DB)}</database>
               </options>),
 "Step 4: load-pipelines(); Status Change Handling, Open XML Extract, and WordprocessinML Process pipelines loaded"
}catch($e){
 $e
}

};

declare function config:attach-pipelines()
{
try{
xdmp:eval("declare namespace config='http://marklogic.com/toolkit/word/cpf-config';
           import module namespace dom = 'http://marklogic.com/cpf/domains' 
		  at '/MarkLogic/cpf/domains.xqy';
           import module namespace p = 'http://marklogic.com/cpf/pipelines' at '/MarkLogic/cpf/pipelines.xqy';
           declare variable $config:d-name as xs:string external;
              (dom:add-pipeline( $config:d-name, 
		  p:get('Status Change Handling')/p:pipeline-id ),
               dom:add-pipeline( $config:d-name, 
		  p:get('Office OpenXML Extract')/p:pipeline-id ),
               dom:add-pipeline( $config:d-name, 
		  p:get('WordprocessingML Process')/p:pipeline-id ))",
               ((xs:QName("config:d-name"), $config:DOMAIN-NAME)),
               <options xmlns="xdmp:eval">
		    <database>{xdmp:database($config:TRIGGERS-DB)}</database>
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

