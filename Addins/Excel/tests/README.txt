Simple unit testing provided:

xqy/excel-unit-test.xqy - tests the functions found in spreadsheet-ml-support.xqy. The tests save to a text file. (edit the path/name as required)  You can set this up to run regularly, doing a diff on the output compared with a base file, so any changes/errors that may occur as you edit the apis are detected.

addin/MarkLogic_ExcelAddin_Test - the source for MarkLogic_ExcelAddin_Test.exe
a c# application that :

1) starts excel
2) saves a workbook as test.xlsx
3) waits 20 seconds
4) saves the workbook
5) closes the workbook
6) exits the excel application


addin/unitTestAddin - 
contains:

unitTestAddin/testMsi.bat
  1) configures the .msi using config.idt
  2) installs the .msi
  3) executes MarkLogic_ExcelAddin_Test.exe
  4) uninstalls the .msi

You need to place the MarkLogic_ExcelAddin_Setup.msi in this directory, as well as config.idt, configured for the application server your Addin is configured with.

add the following: unitTestAddin/MarkLogic_ExcelAddin_Setup.msi
                   unitTestAddin/config.idt


unitTestAddin/PlaceInSamplesDir
the idea is to just leverage the Samples dir you've already configured for Addin development.  Just replace Samples/default.xqy with the default.xqy found here and place test.js under Samples/js/test.js.

The paths are all configured to use C:\unitTestAddin, so if that works for you, you can just copy this directory to that location and execute the testMsi.bat accordingly.

Running the .bat, the .exe that starts the Excel application will run.
When Excel starts, the javascript will run tests onload and write the results out to a file.  The .exe also saves the document created as test.xlsx.

After the test runs, test.xlsx and testresults.txt can be found in unitTestAddin\outputs.








