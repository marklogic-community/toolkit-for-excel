Following are quick up-and-running instructions for those who have already installed the MarkLogic Toolkit for PowerPoint and deployed the Sample application included with the Toolkit.

Deploying the MarkLogic Authoring Sample App for PowerPoint is just as simple,but once deployed into your App Server, 3 areas require update for the application to work properly.

To get started using the Sample Application:

1) Set the registry entry for the Addin to the URL of this application

The key to update is: HKEY_CURRENT_USER/MarkLogicAddinConfiguration/PowerPoint/URL

2) Update <Application-Root>\Author\js\authoring.js

Set the SERVER variable to the URL for the application

3) Update <Application-Root>\Author\config\config.xqy 

Set $config:CONFIG-PATH to the URL for the config.xqy in your application
Set $config:USER to the username for you application server
Set $config:PWD to the password for your application server


NOTE: The Authoring Developer's Guide (pptAuthoringGuide.docx), provides more detail on how to update these files so less configuration may be required for future deployments, as well as Security items to consider that might prompt you to update the code to meet your specific requirements.