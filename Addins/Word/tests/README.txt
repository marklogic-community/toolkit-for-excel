Start of tests for Addin platform.  Following is discussion of contents 
of this directory:

AutomatedTest
+++++++++++++++++++++++++++++++++++++
1)MarkLogic_WordAddin_Setup.msi  
      the Addin installer, you'll want this from cvs, provided here for
      as an example only

2)config.idt
      a config file for updating the URL/properties used by the Addin.
      you'll want this from cvs as well

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
 
5) outputs directory
       Test.docx  - the .docx saved as part of the test
       testfile.txt - the output of the tests
          currently, the javascript writes this file,
          you may want to save as XML to the server

       the contents of these files have currently contain  
       correct output for the tests applied.

MarkLogic_WordAddin_Test
+++++++++++++++++++++++++++++++++++++
the source for the .exe (C#)

addinSampleQATest
+++++++++++++++++++++++++++++++++++++
test.xqy is the page loaded by the Addin on startup
test.js contains the tests

place this directory in the server.  Before running testMsi.bat, 
provide the url for where test.xqy can be found in config.idt.
