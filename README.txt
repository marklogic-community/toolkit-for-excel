Word Authoring Kit Add-In
============================================================
1. MarkLogic-WordAddin



Prerequisites
============================================================
MarkLogic Server:
 version 3.2 or greater


Windows Client: 
 Office 2007 Installed   
    ( 2007 Microsoft Office Primary Interop Assemblies: installed with Office 2007,
      also available separately. )
 .net Framework 3.5
 Visual Studio Tools For Office Runtime 3.0
 Windows Installer 3.1 



Files Required:
============================================================
MarkLogic Server:

MarkLogic-Addin.js
word-processing-ml.xqy
package.xqy


Windows Client:


setup.exe
MarkLogic_WordAddin_Setup.msi

Additional Files:
============================================================
/Samples - directory includes sample Addin examples
    (more detail provided below)

/docs - directory - simple api documentation for the javascript 
                    functions available for interacting with
                    the Active Document in Word.

Notes/Options on Installation For Developers:
===========================================================

If the prerequisites are already installed on the client, only the .msi is
required for installation of the Addin.

setup.exe validates the prereqs on the client, and if not available, prompts
to download and install them from the vendor.  Once the prereqs are installed, 
the .msi will be executed and installed as well.

Configurations for the Addin are stored in the registry at 
HKEY_CURRENT_USER/MarkLogicAddinConfiguration/Word.
These entries will be removed automatically when the application is 
uninstalled.

The .msi installs with these registry entries set to defaults. To edit, you ma
use the Windows SDK to update the msi so you may deliver your solution to users
 with all parameters pre-configured.

The config.idt file is provided with default values.  Just update the values in
the .idt, then execute:

MsiDb -f "<directory where idt is located>" -d "MarkLogic_WordAddin_Setup.msi" -i config.idt

Example:
  MsiDb -f "C:\MyAddin\MyConfig" -d "C:\MyAddin\MarkLogic_WordAddin_Setup.msi" -i config.idt

 ( Note: You'll require the Windows SDK to execute this update.)

This will update the msi with your new values. Another option is to edit these
values using orca.exe ( also available with the Windows SDK ) or just edit the
values in regedit.  Once the .msi has been updated, you may install with
setup.exe, or by itself.

The following registry key values will help to configure your Addin application:

HKEY_CURRENT_USER\MarkLogicAddinConfiguration\Word\
  URL:       The url for the Addin to connect to when the Addin enabled in Word
  RbnBtnLbl: The ribbon Button label
  RbnGrpLbl: The ribbon Group label
  RbnTabLbl: The ribbon Tab label
  CTPTitle:  The title for the Custom Task Pane that has the browser embedded


No other registry values should be edited.



Directions for Installation:
============================================================

MarkLogic Server:

1) Copy word-processing-ml.xqy, package.xqy to 
          <MARKLOGIC>/Modules/MarkLogic/openxml
   (if you are using 4.0, the openxml directory will exist; 
    otherwise, you may need to add it.)
2) Copy MarkLogic-WordAddin.js to whichever directory you will be
   creating your solution in.


Windows Client:
1) copy the addin.deploy folder to your client.

2) Configuration info is stored in the registry. 
   A Key for the current user will be created: 

   HKEY_CURRENT_USER\MarkLogicAddinConfiguration\Word

for this Key, a subkey "URL", contains the value of the url used by
the webBrowser in the Addin when it first loads. 
See the Notes/Option on installation above for information on how to
update these entries. 
 

3) In addin.deploy, Double-Click setup.exe OR MarkLogic_WordAddin_Setup.msi

   If you run setup.exe, the prereqs will be validated on your client. If they
don't exist, you'll be prompted to download and install them from the vendor. 
Once the prereqs are installed, setup.exe executes the .msi to install the Addin.

   If you run MarkLogic_WordAddin_Setup.msi, the Addin will install under the
assumption that all prereqs have been installed prior. There is no validation
of prerequisites.


Usage
============================================================
Upon successful installation of the Add-In, launch Word.  

Samples
============================================================
a /Samples directory is included which provides examples of how to use
the javascript api to interact with the ActiveDocument in Word 
from within the Addin.

To view the samples from within Word, just create an HTTP server in
the MarkLogic Admin interface and set your root directory to
the /Samples dir.
example: /tmp/Samples

set the URL key for the Addin to the new server; 
example (assuming we've created an HTTP server at port 9000
who's root directory is /tmp/Samples):  http://localhost:9000

Start Word and navigate to the Addin.  The landing page provides links for
the examples provided.  Examples of search and reuse, as well as a way
to add custom metadata to a Word document are provided.

Note: a copy of MarkLogic_WordAddin.js is placed in /Samples/js.  

Uninstall
===========================================================
Control Panel -> Add/Remove Programs -> MarkLogic_WordAddin -> Remove
   This will remove the registry entries for the Addin configuration.

Additionally, remove .xqy and .js support sor the Addin from the server.


Troubleshooting
============================================================


Known Issues
============================================================



