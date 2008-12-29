#
#Copyright 2008 Mark Logic Corporation
#
#Licensed under the Apache License, Version 2.0 (the "License");
#you may not use this file except in compliance with the License.
#You may obtain a copy of the License at
#
#    http://www.apache.org/licenses/LICENSE-2.0
#
#Unless required by applicable law or agreed to in writing, software
#distributed under the License is distributed on an "AS IS" BASIS,
#WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
#ee the License for the specific language governing permissions and
#limitations under the License.
#
# Makefile
#
# Instructions: 'make package' compiles Addin, zips up necessary files for Addin, XQuery, javascript, documentation
#  
#

MAJ_VER := `cat MAJOR_VERSION`
MIN_VER := `cat MINOR_VERSION`
DATE := `date +%Y%m%d`
SUFFIX := $(MAJ_VER).$(MIN_VER)-$(DATE)
#ZIP_PREFIX = MarkLogic_WordAddin
ZIP_PREFIX = MarkLogic-Toolkit-for-Word

# Build machine path to MS compiler
# Optional developer machine path to MS compiler
#MS_IDE="C:/Program Files/Microsoft Visual Studio 9.0/Common7/IDE/devenv.exe"
MS_IDE="C:/Program Files (x86)/Microsoft Visual Studio 9.0/Common7/IDE/devenv.exe"
#MS_IDE="C:/WINDOWS/Microsoft.NET/Framework/v3.5/MSBuild.exe"

ML = Addins/Word/xquery
MS = Addins/Word/Microsoft
MSS = MarkLogic_WordAddin
JS = Addins/Word/javascript
XQY = Addins/Word/xquery
CF = Addins/Word/config
SAMPLES = Addins/Word/Samples
DOCS = Addins/Word/docs
JSDOCS = $(DOCS)/jsdocs
SAMPLES_JS = $(SAMPLES)/js

BUILDS = builds
#PUB_BUILD = $(BUILDS)/$(ZIP_PREFIX)-$(SUFFIX)
PUB_BUILD = $(ZIP_PREFIX)-$(SUFFIX)
#PUB_BUILD = $(BUILDS)/Word
ZIP_FILE = $(ZIP_PREFIX)-$(SUFFIX)

BUILD_DOCS = $(PUB_BUILD)/docs
BUILD_DOCS_JSDOC = $(BUILD_DOCS)/jsdocs
BUILD_JS = $(PUB_BUILD)/js
BUILD_XQY =  $(PUB_BUILD)/xquery

BUILD_SAMPLES = $(PUB_BUILD)/Samples
BUILD_SAMPLES_JS = $(PUB_BUILD)/Samples/js
BUILD_SAMPLES_CSS = $(PUB_BUILD)/Samples/css
BUILD_SAMPLES_IMG = $(PUB_BUILD)/Samples/img
BUILD_SAMPLES_METADATA = $(PUB_BUILD)/Samples/metadata
BUILD_SAMPLES_MODULES = $(PUB_BUILD)/Samples/modules
BUILD_SAMPLES_SEARCH = $(PUB_BUILD)/Samples/search

MS_PUB_BUILD = $(PUB_BUILD)/addin.deploy
MS_ROOT = $(MS)/MarkLogic_WordAddin
MS_MAIN_REL = $(MS)/MarkLogic_WordAddin/MarkLogic_WordAddin/bin/Release
MS_MLC_DIR = $(MS_ROOT)/MarkLogic_WordAddin_Setup
MS_BUILD = $(MS_MLC_DIR)/Release
TEMP = temp
#
# Microsoft build (not using MSBuild however, have to use devenv.exe for setup project)
#
# Build machine path to MS compiler
#MS_IDE="C:/Program\ Files/Microsoft\ Visual\ Studio\ 9.0/Common7/IDE/devenv.exe"
#$(MS_IDE) $(MS)/$(MSS)/MarkLogic_WordAddin.sln /build Debug /Out $(MS_LOGS)/debug.log

package: 
	@echo $(MS)/$(MSS)
	@echo $(MS_LOGS)
	@echo $(MS_IDE)
	mkdir $(TEMP)
	mkdir $(BUILDS) 
	mkdir $(PUB_BUILD)
	mkdir $(BUILD_DOCS)
	mkdir $(BUILD_JS)
	mkdir $(BUILD_XQY)
	mkdir $(BUILD_DOCS_JSDOC)
	mkdir $(BUILD_SAMPLES)
	mkdir $(BUILD_SAMPLES_JS)
	mkdir $(BUILD_SAMPLES_CSS)
	mkdir $(BUILD_SAMPLES_IMG)
	mkdir $(BUILD_SAMPLES_METADATA)
	mkdir $(BUILD_SAMPLES_MODULES)
	mkdir $(BUILD_SAMPLES_SEARCH)
	mkdir $(PUB_BUILD)/config
	mkdir $(MS_PUB_BUILD)
	cp $(CF)/*.idt $(PUB_BUILD)/config/.
	cp README.txt $(PUB_BUILD)
	cp  $(MS_ROOT)/$(MSS)/UserControl1.cs  $(TEMP)/UserControl1.cs.bak
	./setVersion patch $(MS_ROOT)/$(MSS)/UserControl1.cs  $(MS_ROOT)/$(MSS)/UserControl2.cs
	rm $(MS_ROOT)/$(MSS)/UserControl1.cs
	mv $(MS_ROOT)/$(MSS)/UserControl2.cs $(MS_ROOT)/$(MSS)/UserControl1.cs
	#build-addin.bat	
	###$(MS_IDE) $(MS)/$(MSS)/MarkLogic_WordAddin.sln /build "Release" /project MarkLogic_WordAddin_Setup/MarkLogic_WordAddin_Setup.vdproj
	#devenv SolutionName /build SolnConfigName [/project ProjName [/projectconfig ProjConfigName]]
	#$(MS_IDE) $(MS_MLC_DIR)/MarkLogic_WordAddin_Setup.vdproj /build Release 
	$(MS_IDE) $(MS_MLC_DIR)/MarkLogic_WordAddin_Setup.vdproj /build "Release"
	mv $(TEMP)/UserControl1.cs.bak  $(MS_ROOT)/$(MSS)/UserControl1.cs
	#here
	cp -r   $(MS_BUILD)/* $(MS_PUB_BUILD)/.
	./setVersion patch $(JS)/MarkLogicWordAddin.js $(BUILD_JS)/MarkLogicWordAddin.js
	./setVersion patch $(JS)/MarkLogicWordAddin.js $(SAMPLES_JS)/MarkLogicWordAddin.js
	#./setVersion patch $(ML)/word-processing-ml.xqy $(PUB_BUILD)/word-processing-ml.xqy
	#./setVersion patch $(ML)/package.xqy $(PUB_BUILD)/package.xqy
	#cp -r $(SAMPLES)/* $(BUILD_SAMPLES) 
	cp $(XQY)/word-processing-ml-support.xqy $(BUILD_XQY)
	cp $(SAMPLES)/default.xqy $(BUILD_SAMPLES)
	cp $(SAMPLES)/README.txt $(BUILD_SAMPLES)
	cp $(SAMPLES)/samples-license.txt $(BUILD_SAMPLES)
	cp $(SAMPLES)/js/*.js $(BUILD_SAMPLES_JS) 
	cp $(SAMPLES)/css/*.css $(BUILD_SAMPLES_CSS) 
	cp $(SAMPLES)/img/*.png $(BUILD_SAMPLES_IMG) 
	cp $(SAMPLES)/img/LICENSE $(BUILD_SAMPLES_IMG) 
	cp $(SAMPLES)/metadata/*.js $(BUILD_SAMPLES_METADATA) 
	cp $(SAMPLES)/metadata/*.xqy $(BUILD_SAMPLES_METADATA) 
	cp $(SAMPLES)/metadata/*.css $(BUILD_SAMPLES_METADATA) 
	cp $(SAMPLES)/search/*.js $(BUILD_SAMPLES_SEARCH) 
	cp $(SAMPLES)/search/*.xqy $(BUILD_SAMPLES_SEARCH) 
	cp $(SAMPLES)/search/*.css $(BUILD_SAMPLES_SEARCH) 
	cp $(JSDOCS)/*.css $(BUILD_DOCS_JSDOC)
	cp $(JSDOCS)/*.html $(BUILD_DOCS_JSDOC)
	cp $(DOCS)/ToolkitForWordGuide.pdf $(BUILD_DOCS)
	#cp -r $(SAMPLES)/modules/*.js $(BUILD_SAMPLES_JS) 
	@echo Create zip file $(ZIP_PREFIX)_$(SUFFIX).zip
	#(cd builds; zip -r ../$(ZIP_FILE).zip $(PUB_BUILD)/*)
	zip -r $(ZIP_FILE).zip $(PUB_BUILD)/*
	mv $(PUB_BUILD) $(BUILDS)
	#mv $(ZIP_PREFIX).zip $(ZIP_PREFIX)-$(SUFFIX).zip

clean:
	  rm -rf $(BUILDS)
	  rm -rf $(MS_BUILD)/*
	  rm -rf $(MS_MAIN_REL)/*
	  rm -rf $(TEMP)
	  rm ./*.zip

