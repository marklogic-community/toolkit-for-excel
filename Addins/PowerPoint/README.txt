MarkLogic Toolkit for PowerPoint®

MarkLogic Add-in for PowerPoint®

The MarkLogic Toolkit for PowerPoint® allows you to integrate Microsoft PowerPoint
with MarkLogic Server.

The ToolkitForPowerPointGuide.docx document in the /docs/ directory of the zip package
contains the documentation for the MarkLogic Toolkit for PowerPoint®, and includes 
information on system requirements, installation of the 
MarkLogic Add-in for PowerPoint®, and configuration of the installer program
to deploy a customized installer to your Microsof Word user base.  
The latest version of the documention is available on 
http://developer.marklogic.com/pubs.

Copyright 2002-2010 Mark Logic Corporation.  All Rights Reserved.

Change Notes:
------------------
1.0-2 update of presentation-ml-support.xqy.  
      fixed mapping of slide references and sorting of slides in ppt:insert-slide() function

1.0-3 update of presenation-ml-support.xqy.  
      fixed ppt:map-max-image-id() to return 1 in case of empty seq.  previously was erroring out if no image in package.

      update of presentation-ml-support.html
      fixed documentation to remove comments and replaced calls ppt:package-map-create() to ppt:package-map()

