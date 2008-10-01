xquery version "0.9-ml"
declare namespace w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"

xdmp:set-response-content-type('text/html'),
      <html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en">
        <head>
            <title>WordAddOn</title>
            <meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
	    <link type="text/css" rel="stylesheet" href="metadata.css" />
            <script type="text/javascript" src="metadata.js"></script>
            <script type="text/javascript" src="../js/MarkLogicWordAddin.js"></script>

            
        </head>
        {
        let $rgb :=  "rgb(200,216,237)"
        let $body:=
         <body bgcolor={$rgb}>
              <div id="header">
                 <form id="basicsearch" action="#" method="post">
                   <div>
                      Title <br/>
                      <input type="text" name="v_title" autocomplete="off" value="" style='width: 250px;'id="v_title"/>&nbsp;
                      <br/>
                      Description <br/>
                      <textarea rows="5" cols="30" name="v_desc" value="" style='width: 250px; height: 90px;' id="v_desc"/>&nbsp;
                      <br/>
                      Publisher <br/>
                      <input type="text" name="v_publisher" value="" style='width: 250px;' id="v_pub"/>
                      <br/>
                      Identifier <br/>
                      <input type="text" name="v_identifier" value="" style='width: 250px;' id="v_id"/>
                      <br/>
                      <br/>
                      <input type="submit" value="Save Metadata" onClick="updateMetadata(1)"/>&nbsp;&nbsp;
                      <input type="submit" value="Remove" style='width: 110px;'  onClick="updateMetadata(2)"/>
                     <br/>
                     <br/>
                     <p style="color:#000000;" id="v_fc"></p>    
                     <br/>
                     <p style="color:#000000;" id="v_fc2"></p>    
                                             
                   </div> 
                  </form>
                   
             </div>
          </body>
          return $body
         }
         </html>
