rem cd C:\cygwin\home\paven\MLOS\Office2007\Addins\Word\Microsoft\MarkLogic_WordAddin\MarkLogic_WordAddin
cd Addins\Word\Microsoft\MarkLogic_WordAddin\MarkLogic_WordAddin

rem call "C:\Program Files (x86)\Microsoft Visual Studio 9.0\VC\vcvarsall.bat"
call  "C:\Program Files\Microsoft Visual Studio 9.0\VC\vcvarsall.bat"

msbuild /target:publish
