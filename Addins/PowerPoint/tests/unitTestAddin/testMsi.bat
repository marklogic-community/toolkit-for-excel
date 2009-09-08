REM This line updates the .msi, by importing config.idt into the .msi.
REM The command is MsiDb
REM options:
REM -f folder where .idt is found
REM -d the .msi to be updated
REM -i import the .idt
REM (when these 3 arguments are passed, the command executes silently,without requesting user input)
"C:\Program Files\Microsoft SDKs\Windows\v6.0A\bin\MsiDb" -f "C:\unitTestAddin" -d "MarkLogic_PowerPointAddin_Setup.msi" -i config.idt"


REM installs the .msi
REM msiexec should be in your path if you have Windows Installer installed (Addin requires 3.1 or greater)
REM options:
REM /q - quiet
REM /i - install
REM TARGETDIR - directory to install msi to
msiexec /q /i "C:\unitTestAddin\MarkLogic_PowerPointAddin_Setup.msi" TARGETDIR="C:\tmp"

REM executes the C# app
REM this will start PowerPoint, and close PowerPoint after running tests specifed by parameter
REM output file of test results writtent to C:\unitTestAdin\outputs\TestResults.txt
REM the tests are run from the javascript in the page when it loads in the browswer
MarkLogic_PowerPointAddin_Test.exe 1  

REM uninstalls the Addin
msiexec /q /x "C:\unitTestAddin\MarkLogic_PowerPointAddin_Setup.msi" 


