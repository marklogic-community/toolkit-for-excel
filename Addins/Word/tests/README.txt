Start of tests for Addin platform.  Following is discussion of contents 
of this directory:

AutomatedTest
+++++++++++++++++++++++++++++++++++++
1)MarkLogic_WordAddin_Setup.msi  
      the Addin installer, you'll want this from cvs/svn, provided here for
      as an example only

2)config.idt
      a config file for updating the URL/properties used by the Addin.
      you'll want this from cvs/svn as well

3)MarkLogic_WordAddin_Test.exe
      a C# program that starts Word, saves the .docx, stays open for 
      20 seconds, then closes Word.

4)testMsi.bat
      comments available in . bat file
      this file:
       a) updates .msi using config.idt
       b) installs the .msi
       c) executes MarkLogic_WordAddin_Test.exe
       d) uninstalls the .msi
 
5) testInput directory
       seed .docx files for various tests
       in testMsi.bat you can see the file will be opened for testing, 
       and the result will be output to /testOutput directory

6) testOutput
       the contents of these files currently contain  
       correct output for the tests applied.


MarkLogic_WordAddin_Test
+++++++++++++++++++++++++++++++++++++
the source for the .exe (C#)

wordQATests (aka PlaceInSamplesDir)
+++++++++++++++++++++++++++++++++++++
test.xqy is the page loaded by the Addin on startup
test.js contains the tests
MarkLogicWordAddin.js - add this from cvs/svn

other .xqy files used for upload/download tests
.css and .png provided for simple styling

place this directory in the server.  Before running testMsi.bat, 
provide the url for where test.xqy can be found in config.idt.  

When the Addin is loaded, and the browser starts up, the tests run from test.js onLoad().  
outputs are written to testOutput/<filename>.txt, and the .docx is saved to /testOutput as well.
