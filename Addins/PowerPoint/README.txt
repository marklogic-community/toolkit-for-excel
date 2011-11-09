MarkLogic Toolkit for PowerPoint®

MarkLogic Add-in for PowerPoint®

The MarkLogic Toolkit for PowerPoint® allows you to integrate Microsoft PowerPoint
with MarkLogic Server.

The ToolkitForPowerPointGuide.docx document in the /docs/ directory of the zip package
contains the documentation for the MarkLogic Toolkit for PowerPoint®, and includes 
information on system requirements, installation of the 
MarkLogic Add-in for PowerPoint®, and configuration of the installer program
to deploy a customized installer to your Microsof Word user base.  
The latest version of the documention is available on 
http://developer.marklogic.com/pubs.

Copyright 2002-2010 Mark Logic Corporation.  All Rights Reserved.

Change Notes:
------------------
1.0-2 update of presentation-ml-support.xqy.  
      fixed mapping of slide references and sorting of slides in ppt:insert-slide() function

1.0-3 update of presenation-ml-support.xqy.  
      fixed ppt:map-max-image-id() to return 1 in case of empty seq.  previously was erroring out if no image in package.

      update of presentation-ml-support.html
      fixed documentation to remove comments and replaced calls ppt:package-map-create() to ppt:package-map()

1.1-1 update of MarkLogicPowerPointAddin.js - 27 new functions 
       
      New Functions
      =============
      getSlideName()			get name of active slide    
      getSlideIndex()			get index of active slide
      getPresentationSlideCount()	count for all slides in presentation
      addSlideTag()			add tag to slide by index     
      deleteSlideTag()		        delete tag from slide by index, slidetag
      getSlideTags()			get all tags for slide by index
      addShapeTag()			add tag to shape by shapename
      deleteShapeTag()		        delete tag from shape by shapename, tagname
      getShapeRangeName()	   	get shapename of selected shape
      getShapeRangeShapeNames() 	return all selected shape names
      addShapeRangeTag()        	add tags to all selected shapes
      addPresentationTag()		add tags to active presentation
      deletePresentationTag()	        delete tags from active presentation
      getPresentationTags() 		get tags for active presentation
      getShapeRangeCount()      	count of selected shapes in current slide
      setShapeRangeName()		set name of shape to override MS default
      getShapeRangeView() 		returns shape info as MLA.ShapeRangeView object 
      jsonStringify()			serializes JS object as JSON string
      jsonParse()			takes JSON serialization and constructs JS object
      addShape()			takes slide index and MLA.ShapeRangeView object as 
      setPictureFormat()	//takes picture format (as JSON) and applies to picture shape identified by name
      addPresentationTags()	//takes JSON string of tags and sets to active presentation
      addSlideTags()		//takes JSON string of tags and sets to slide identified by index
      addShapeTags()    	//takes JSON string of tags and sets to shape identified by name
      addSlide()
      deleteSlide()
      deleteShape()

      added MarkLogicPowerPointEventSupport.js - 19 events captured

      Events
      =======
      windowSelectionChange()
      windowBeforeRightClick()
      windowBeforeDoubleClick()
      presentationClose()
      presentationSave()
      presentationOpen()
      newPresentation()
      presentationNewSlide()
      windowActivate()
      windowDeactivate()
      slideShowBegin()
      slideShowNextBuild()
      slideShowNextSlide()
      slideShowEnd()
      presentationPrint()
      slideSelectionChanged()
      colorSchemeChanged()
      presentationBeforeSave()
      slideShowNextClick()


      New CPF Pipeline PresentationML Tags Process
      Pipeline sets tags for presentations, slides, and slide components 
       on document properties of presentation.xml and slideN.xml

      added presentationml-tags-pipeline.xml - pipeline definition
      added pptx-set-tags-action.xqy         - pipeline action
      updated presentationml-pipeline.xml    - now initial state is tagged, instead of initial
      update of presentation-ml-support.xqy (unpublished function to assist dereferencing of tags)
 
2.0 install/install.xqy script provided to simplify install of .xqy and CPF components 
