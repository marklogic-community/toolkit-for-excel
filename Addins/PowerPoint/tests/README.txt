Simple unit testing provided:

xqy/ppt-unit-test.xqy :
====================================
Tests the functions found in presentation-ml-support.xqy. (edit the path/name as required)  You can set this up to run regularly, doing a diff on the output compared with a base file, so any changes/errors that may occur as you edit the apis are detected.

The .xqy tests are dependant on presentations saved and unzipped in the database.  Update accordingly.   PowerPoint files are saved to /unitTestAddin/pptx, other files can be found in /unitTestAddin/outputs.

map.xml : output from ppt:package-map()

testOne.xml : output from ppt:package-make()



/MarkLogic_PowerPointAddin_Test :
====================================
The source for MarkLogic_PowerPointAddin_Test.exe
a c# application that :

1) starts PowerPoint application

and depending on Option:

option 1:
 A) loops through a directory of .pptx, opening and closing each one.  These .pptx should be modified/generated using presentation-ml-support.xqy
 B) Quits PowerPoint application

 Results are written to: /unitTestAddin/outputs/TestResults.txt


option 2:
  A) Adds a presentation with a default slide.
  B) sleeps for 20 seconds
  C) closes presentation
  C) Quits PowerPoint application

  The assumption for option 2 is that testing is performed using the onLoad of test.js so addin libraries can be called to modify the presentation.


/unitTestAddin/testMsi.bat :
====================================

  1) configures the .msi using config.idt
  2) installs the .msi
  3) executes MarkLogic_PowerPointAddin_Test.exe
  4) uninstalls the .msi

You need to place the MarkLogic_PowerPointAddin_Setup.msi in this same directory, as well as config.idt, configured for the application server your Addin is configured with.

add the following: unitTestAddin/MarkLogic_PowerPointAddin_Setup.msi
                   unitTestAddin/config.idt


/unitTestAddin/PlaceInSamplesDir :
===================================
the idea is to just leverage the Samples dir you've already configured for Addin development.  Just replace Samples/default.xqy with the default.xqy found here and place test.js under Samples/test.js.

The paths are all configured to use C:\unitTestAddin, so if that works for you, you can just copy this directory to that location and execute the testMsi.bat accordingly.

Running the .bat, the .exe that starts the PowerPoint application will run.
When PowerPoint starts, the javascript will run tests onload and write the results out to a file in /unitTestAddin/outputs

Results of test are written to: /unitTestAddin/outputs.










