Word Authoring Kit Add-In
============================================================
1. MarkLogic-WordAddin



Prerequisites
============================================================
MarkLogic Server:
 version 3.2 or greater


Windows Client:
 Office 2007 Installed
 .net Framework 3.5
 Visual Studio Tools For Office Runtime 3.0



Files Required:
============================================================
MarkLogic Server:

MarkLogic-Addin.js
word-processing-ml.xqy
package.xqy


Windows Client:

ntConfigAdd.reg

setup.exe
MarkLogic-OfficeSuite.vsto
/Application Files (directory)



Directions for Installation:
============================================================

MarkLogic Server:

1) Copy word-processing-ml.xqy, package.xqy to <MARKLOGIC>/Modules/MarkLogic/openxml
   (if you are using 4.0, the openxml directory will exist; otherwise, you may need to add it.)
2) Copy MarkLogic-WordAddin.js to whichever directory you will be creating your solution in.


Windows Client:
1) copy the addin.deploy folder as well as the .reg file to your client.

2) Configuration info is stored in the registry. A Key for the current user will be created: 

HKEY_CURRENT_USER\MarkLogicAddinConfiguration\Word

for this Key, a subkey "URL", contains the value of the url used by the webBrowser in the Addin when it first loads.  
  A)Update the value in ntConfigAdd.reg to be the desired url.
  B)Double click the .reg file, and the entry will be added to your system.

To remove the Key (and it's values), double click ntConfigRemove.reg.

3) In addin.deploy, Double-Click setup.exe (if installing on Vista, please right-click, and install as Administrator).

   The Add-In will install.  It requires .net 3.5 and the Visual Studio Tools for Office Runtime 3.0.  If these aren't already available, the prerequisites will be downloaded from Microsoft.  You'll be prompted to install these as well.



Usage
============================================================
Upon successful installation of the Add-In, launch Word.  




Uninstall
===========================================================
Control Panel -> Add/Remove Programs -> MarkLogic-WordAddin -> Remove

double click ntConfigRemove.reg to remove the registry entry.

Additionally, remove .xqy and .js support sor the Addin from the server.


Troubleshooting
============================================================


Known Issues
============================================================



