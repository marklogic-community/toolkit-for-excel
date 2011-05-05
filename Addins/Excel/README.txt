MarkLogic Toolkit for Excel®

MarkLogic Add-in for Excel®

The MarkLogic Toolkit for Excel® allows you to integrate Microsoft Excel
with MarkLogic Server.

The ToolkitForExcelGuide.docx document in the docs/ directory of the zip package
contains the documentation for the MarkLogic Toolkit for Excel®, and includes 
information on system requirements, installation of the 
MarkLogic Add-in for Excel®, and configuration of the installer program
to deploy a customized installer to your Microsof Excel user base.  
The latest version of the documention is available on 
http://developer.marklogic.com/pubs.

Copyright 2002-2011 Mark Logic Corporation.  All Rights Reserved.

Change Notes:
------------------
version 2.0

new functions:

MLA.addChartObjectMouseDownEvents()
MLA.addMacro()
MLA.base64EncodeImage()
MLA.deleteFile()
MLA.deletePicture()
MLA.exportChartImagePNG()
MLA.getMacroComments()
MLA.getMacroCount()
MLA.getMacroName()
MLA.getMacroProcedureName()
MLA.getMacroSignature()
MLA.getMacroText()
MLA.getMacroType()
MLA.getSelectedChartName()
MLA.getSelectedRangeName()
MLA.getSheetType()
MLA.getWorksheetChartNames()
MLA.getWorksheetNamedRangeNames()
MLA.insertBase64ToImage()
MLA.removeChartObjectMouseDownEvents()
MLA.removeMacro()
MLA.runMacro()

now capture following events:

sheetActivate()
sheetBeforeDoubleClick()
sheetBeforeRightClick()
sheetChange()
sheetDeactivate()
rangeSelected() 
    -sheetSelectionChange Event only caught when selection is named range
workbookActivate()
workbookAfterXmlExport()
workbookAfterXmlImport()
workbookBeforeXmlExport()
workbookBeforeXmlImport()
workbookBeforeClose()
workbookBeforeSave()
workbookDeactivate()
workbookNewSheet()
workbookOpen()
chartObjectMouseDown()

custom event definitions should be placed by developers in:
   MarkLogicExcelEventHandlers.js

