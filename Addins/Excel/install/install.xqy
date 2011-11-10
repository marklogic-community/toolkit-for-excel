xquery version "1.0-ml";
(: Copyright 2009-2011 MarkLogic Corporation

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
:)

declare namespace config="http://marklogic.com/toolkit/excel/cpf-config";
import module namespace dom = "http://marklogic.com/cpf/domains" 
       at "/MarkLogic/cpf/domains.xqy";
import module namespace p = "http://marklogic.com/cpf/pipelines"
       at "/MarkLogic/cpf/pipelines.xqy";

(: Source directory for spreadsheet-ml-support.xqy :)
declare variable $config:SUPPORT-SRC-PATH := 
          "C:/Users/user/Desktop/Excel/xquery/";

(: Source directory for map-shared-action.xqy :)
declare variable $config:CPF-SRC-PATH :=
          "C:/Users/user/Desktop/Excel/cpf/";

(: Support .xqy files will be copied from $config:SUPPORT-SRC-PATH
   and placed under $config:SERVER-ROOT/Modules/MarkLogic/openxml/".

   CPF .xqy file will be copied from $config:CPF-SRC-PATH 
   and placed under $config:SERVER-ROOT/Modules/MarkLogic/conversion/actions/"

   CPF .xml file will be copied from $config:CPF-SRC-PATH 
   and placed under $config:SERVER-ROOT/Installer/openxml

   If you have a non-standard configuration of MarkLogic Server,  
   update $config:SERVER-ROOT accordingly. :)
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

declare variable $config:TRIGGERS-DB := "Triggers";

(:  If you already have CPF installed with a Domain configured 
    for the same $config:DOMAIN-URI you'll end up with 2 domains
    with different names, but for the same URI. 
    Don't cross the streams!  Rename $DOMAIN-NAME accordingly. :)
declare variable $config:DOMAIN-NAME := "OpenXML";
declare variable $config:DOMAIN-DESCRIPTION := 
                     "Handling incoming Open XML documents";
declare variable $config:DOMAIN-SCOPE := "directory";
declare variable $config:DOMAIN-URI := "/";
declare variable $config:DOMAIN-DEPTH := "infinity";
declare variable $config:CONTEXT-DB := "Modules";
declare variable $config:CONTEXT-ROOT :=     "/";

declare function config:install-xqy-support()
{
 try{
  let $support-file := "spreadsheet-ml-support.xqy"
  let $cpf-action := "map-shared-action.xqy"
  let $cpf-config := "spreadsheetml-pipeline.xml"

  let $openxml-path := fn:concat($config:SERVER-ROOT,
                                 "Modules/MarkLogic/openxml/")
  let $cpf-path := fn:concat($config:SERVER-ROOT,
                             "Modules/MarkLogic/conversion/actions/")
  let $install-path := fn:concat($config:SERVER-ROOT,
                                 "Installer/openxml/")

  let $src-support  := fn:concat($config:SUPPORT-SRC-PATH, $support-file)
  let $dest-support := fn:concat($openxml-path, $support-file)
  let $support-doc  := xdmp:document-get($src-support)
  let $src-cpf-action := fn:concat($config:CPF-SRC-PATH,$cpf-action)
  let $src-cpf-config :=  fn:concat($config:CPF-SRC-PATH,$cpf-config)
  let $dest-cpf-action:= fn:concat($cpf-path,$cpf-action)
  let $dest-cpf-install := fn:concat($install-path,$cpf-config)
  let $cpf-action-doc := xdmp:document-get($src-cpf-action)
  let $cpf-config-doc := xdmp:document-get($src-cpf-config)

  return (xdmp:save($dest-support, $support-doc),
          xdmp:save($dest-cpf-action, $cpf-action-doc),
          xdmp:save($dest-cpf-install, $cpf-config-doc),
          fn:concat("Step 1: install-xqy-support(); ",
                    $support-file,
                    " copied to ",
                    $openxml-path),
          fn:concat("Step 1: install-xqy-support(); ",
                    $cpf-action,
                    " copied to ",
                    $cpf-path),
          fn:concat("Step 1: install-xqy-support(); ",
                    $cpf-config,
                    " copied to ",
                    $install-path)
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
      import module namespace p = 'http://marklogic.com/cpf/pipelines' 
             at '/MarkLogic/cpf/pipelines.xqy';
      declare variable $config:d-name as xs:string external;
      declare variable $config:d-description as xs:string external;
      declare variable $config:d-scope as xs:string external;
      declare variable $config:d-URI as xs:string external;
      declare variable $config:d-depth as xs:string external;
      declare variable $config:c-db as xs:string external;
      declare variable $config:c-root as xs:string external;

      dom:create( $config:d-name, 
                  $config:d-description, 
                  dom:domain-scope( $config:d-scope, 
                                    $config:d-URI, 
		                    $config:d-depth),
	          dom:evaluation-context( xdmp:database($config:c-db), 
		                          $config:c-root ),
                  (), 
                  () 
               )", 
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
     if(fn:string($e/error:code) eq "CPF-DOMAINEXISTS") then
        "Step 2: create-domain() encountered the following: 
         CPF-DOMAINEXISTS - Domain Already Defined"
     else 
        fn:string($e/error:code)
  }
};

declare function config:create-configuration()
{
  try{
    xdmp:eval("
      declare namespace config='http://marklogic.com/toolkit/excel/cpf-config';
      import module namespace dom = 'http://marklogic.com/cpf/domains' 
	     at '/MarkLogic/cpf/domains.xqy';
      import module namespace p = 'http://marklogic.com/cpf/pipelines' 
             at '/MarkLogic/cpf/pipelines.xqy';
      declare variable $config:r-user as xs:string external;
      declare variable $config:c-db as xs:string external;
      declare variable $config:d-name as xs:string external;

      dom:configuration-create($config:r-user, 
                               dom:evaluation-context( 
                                  xdmp:database($config:c-db), 
                                  '/' 
                               ),
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
     if(fn:string($e/error:code) eq "CPF-CONFIGEXISTS") then
        "Step 3: create-configuration() encountered the following: 
         CPF-CONFIGEXISTS - Configuration Already Exists"
     else 
         fn:string($e/error:code)
  }
};

declare function config:load-pipelines()
{
  try{
    xdmp:eval("
      declare namespace config='http://marklogic.com/toolkit/excel/cpf-config';
      import module namespace dom = 'http://marklogic.com/cpf/domains' 
             at '/MarkLogic/cpf/domains.xqy';
      import module namespace p = 'http://marklogic.com/cpf/pipelines' 
             at '/MarkLogic/cpf/pipelines.xqy';
      declare variable $config:cpf-path as xs:string external;
           
      (p:insert(xdmp:document-get('Installer/cpf/status-pipeline.xml')),
       p:insert(xdmp:document-get('Installer/openxml/openxml-pipeline.xml')),
       p:insert(xdmp:document-get(
                 fn:concat($config:cpf-path,'/spreadsheetml-pipeline.xml')
                )
               )
      )",
      ((xs:QName("config:cpf-path"), fn:concat($config:SERVER-ROOT,
                                               "Installer/openxml/")
      )),
      <options xmlns="xdmp:eval">
	  <database>{xdmp:database($config:TRIGGERS-DB)}</database>
      </options>),
      "Step 4: load-pipelines(); Status Change Handling, Open XML Extract,
       and SpreadsheetML Process pipelines loaded"
  }catch($e){
     $e
  }
};

declare function config:attach-pipelines()
{
  try{
    xdmp:eval("
      declare namespace config='http://marklogic.com/toolkit/excel/cpf-config';
      import module namespace dom = 'http://marklogic.com/cpf/domains' 
             at '/MarkLogic/cpf/domains.xqy';
      import module namespace p = 'http://marklogic.com/cpf/pipelines' 
             at '/MarkLogic/cpf/pipelines.xqy';
      declare variable $config:d-name as xs:string external;
      (dom:add-pipeline($config:d-name, 
		        p:get('Status Change Handling')/p:pipeline-id ),
       dom:add-pipeline($config:d-name,
                        p:get('Office OpenXML Extract')/p:pipeline-id ),
       dom:add-pipeline($config:d-name, 
		        p:get('SpreadsheetML Process')/p:pipeline-id ))",
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

