xquery version "1.0-ml";


declare namespace config="http://marklogic.com/toolkit/ppt/cpf-config";
import module namespace dom = "http://marklogic.com/cpf/domains" 
       at "/MarkLogic/cpf/domains.xqy";
import module namespace p = "http://marklogic.com/cpf/pipelines" 
       at "/MarkLogic/cpf/pipelines.xqy";

(: Source directory for presentation-ml-support.xqy,
                        presentation-ml-support-content-types.xqy :)
declare variable $config:SUPPORT-SRC-PATH := 
          "C:\Users\paven\Desktop\PowerPoint\xquery\";

(: Source directory for map-slide-action.xqy, 
                        pptx-set-tags-action :)
declare variable $config:CPF-SRC-PATH :=
          "C:\Users\paven\Desktop\PowerPoint\cpf\";

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

declare variable $config:TRIGGERS-DB := "TK-Triggers";

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
declare variable $config:CONTEXT-ROOT := "/";

declare function config:install-xqy-support()
{
 try{
  let $support-file-1 := "presentation-ml-support.xqy"
  let $support-file-2 := "presentation-ml-support-content-types.xqy"
  let $cpf-action-1   := "map-slide-action.xqy"
  let $cpf-action-2   := "pptx-set-tags-action.xqy"
  let $cpf-config-1   := "presentationml-pipeline.xml"
  let $cpf-config-2   := "presentationml-tags-pipeline.xml"

  let $openxml-path := fn:concat($config:SERVER-ROOT,
                                 "Modules/MarkLogic/openxml/")
  let $cpf-path :=     fn:concat($config:SERVER-ROOT,
                                 "Modules/MarkLogic/conversion/actions/")
  let $install-path := fn:concat($config:SERVER-ROOT,
                                 "Installer/openxml/")

  let $src-support-1  := fn:concat($config:SUPPORT-SRC-PATH, 
                                   $support-file-1)
  let $src-support-2  := fn:concat($config:SUPPORT-SRC-PATH, 
                                   $support-file-2)
  let $dest-support-1 := fn:concat($openxml-path, 
                                   $support-file-1)
  let $dest-support-2 := fn:concat($openxml-path, 
                                   $support-file-2)
  let $support-doc-1  := xdmp:document-get($src-support-1)
  let $support-doc-2  := xdmp:document-get($src-support-2)

  let $src-cpf-action-1 := fn:concat($config:CPF-SRC-PATH,$cpf-action-1)
  let $src-cpf-action-2 := fn:concat($config:CPF-SRC-PATH,$cpf-action-2)
  let $src-cpf-config-1 :=  fn:concat($config:CPF-SRC-PATH,$cpf-config-1)
  let $src-cpf-config-2 :=  fn:concat($config:CPF-SRC-PATH,$cpf-config-2)
  let $dest-cpf-action-1:= fn:concat($cpf-path,$cpf-action-1)
  let $dest-cpf-action-2:= fn:concat($cpf-path,$cpf-action-2)
  let $dest-cpf-install-1 := fn:concat($install-path,$cpf-config-1)
  let $dest-cpf-install-2 := fn:concat($install-path,$cpf-config-2)
  let $cpf-action-doc-1 := xdmp:document-get($src-cpf-action-1)
  let $cpf-action-doc-2 := xdmp:document-get($src-cpf-action-2)
  let $cpf-config-doc-1 := xdmp:document-get($src-cpf-config-1)
  let $cpf-config-doc-2 := xdmp:document-get($src-cpf-config-2)

  return (xdmp:save($dest-support-1, $support-doc-1),
          xdmp:save($dest-support-2, $support-doc-2),
          xdmp:save($dest-cpf-action-1, $cpf-action-doc-1),
          xdmp:save($dest-cpf-action-2, $cpf-action-doc-2),
          xdmp:save($dest-cpf-install-1, $cpf-config-doc-1),
          xdmp:save($dest-cpf-install-2, $cpf-config-doc-2),
          fn:concat("Step 1: install-xqy-support(); ",
                    $support-file-1,
                    " copied to ",
                    $openxml-path),
          fn:concat("Step 1: install-xqy-support(); ",
                    $support-file-2,
                    " copied to ",
                    $openxml-path),
          fn:concat("Step 1: install-xqy-support(); ",
                    $cpf-action-1,
                    " copied to ",
                    $cpf-path),
          fn:concat("Step 1: install-xqy-support(); ",
                    $cpf-action-2,
                    " copied to ",
                    $cpf-path),
          fn:concat("Step 1: install-xqy-support(); ",
                    $cpf-config-1,
                    " copied to ",
                    $install-path),
          fn:concat("Step 1: install-xqy-support(); ",
                    $cpf-config-2,
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
        declare namespace config='http://marklogic.com/toolkit/ppt/cpf-config';
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
	</options>
     ),
     "Step 2: create-domain(); Domain Created"
 }catch($e){
      if(fn:string($e/error:code) eq "CPF-DOMAINEXISTS")then
        "Step 2: create-domain() encountered the following: 
         CPF-DOMAINEXISTS - Domain Already Defined"
      else fn:string($e/error:code)
 }

};

declare function config:create-configuration()
{
  try{
     xdmp:eval("
        declare namespace config='http://marklogic.com/toolkit/ppt/cpf-config';
        import module namespace dom = 'http://marklogic.com/cpf/domains' 
	       at '/MarkLogic/cpf/domains.xqy';
        import module namespace p = 'http://marklogic.com/cpf/pipelines' 
               at '/MarkLogic/cpf/pipelines.xqy';
        declare variable $config:r-user as xs:string external;
        declare variable $config:c-db as xs:string external;
        declare variable $config:d-name as xs:string external;

        dom:configuration-create( 
                          $config:r-user, 
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
      if(fn:string($e/error:code) eq "CPF-CONFIGEXISTS")then
         "Step 3: create-configuration() encountered the following: 
          CPF-CONFIGEXISTS - Configuration Already Exists"
      else fn:string($e/error:code)
}

};

declare function config:load-pipelines()
{
  try{
    xdmp:eval("
        declare namespace config='http://marklogic.com/toolkit/ppt/cpf-config';
        import module namespace dom = 'http://marklogic.com/cpf/domains' 
	       at '/MarkLogic/cpf/domains.xqy';
        import module namespace p = 'http://marklogic.com/cpf/pipelines' 
               at '/MarkLogic/cpf/pipelines.xqy';
        declare variable $config:cpf-path as xs:string external;
           
        (p:insert(
            xdmp:document-get('Installer/cpf/status-pipeline.xml')
         ),
         p:insert(
            xdmp:document-get('Installer/openxml/openxml-pipeline.xml')
         ),
         p:insert(
            xdmp:document-get(
                  fn:concat($config:cpf-path,'/presentationml-pipeline.xml')
            )
         ),
         p:insert(
            xdmp:document-get(
                  fn:concat($config:cpf-path,'/presentationml-tags-pipeline.xml')
            )
         )
    )",
    (
      (xs:QName("config:cpf-path"), 
       fn:concat($config:SERVER-ROOT,
       "Installer/openxml/"))
    ),
    <options xmlns="xdmp:eval">
        <database>{xdmp:database($config:TRIGGERS-DB)}</database>
    </options>),
    "Step 4: load-pipelines(); Status Change Handling, Open XML Extract, 
     PresentationML Process, and PresentationML Process Tags pipelines loaded"
  }catch($e){
     $e
  }

};

declare function config:attach-pipelines()
{
  try{
    xdmp:eval("
        declare namespace config='http://marklogic.com/toolkit/ppt/cpf-config';
        import module namespace dom = 'http://marklogic.com/cpf/domains' 
	       at '/MarkLogic/cpf/domains.xqy';
        import module namespace p = 'http://marklogic.com/cpf/pipelines' 
               at '/MarkLogic/cpf/pipelines.xqy';
        declare variable $config:d-name as xs:string external;

        (dom:add-pipeline( $config:d-name, 
		           p:get('Status Change Handling')/p:pipeline-id ),
         dom:add-pipeline( $config:d-name, 
		           p:get('Office OpenXML Extract')/p:pipeline-id ),
         dom:add-pipeline( $config:d-name, 
		           p:get('PresentationML Process')/p:pipeline-id ),
         dom:add-pipeline( $config:d-name, 
		           p:get('PresentationML Process Tags')/p:pipeline-id )
        )",
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

