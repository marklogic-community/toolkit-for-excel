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

MAJ_VER :=2 #`cat MAJOR_VERSION`
MIN_VER :=0 #`cat MINOR_VERSION`
DATE := `date +%Y%m%d`
#SUFFIX := $(MAJ_VER).$(MIN_VER)-$(DATE)
SUFFIX := 2.0-1-20111123
#$(DATE)
#ZIP_PREFIX = MarkLogic_WordAddin
ZIP_PREFIX = MarkLogic-Toolkit-for-Excel

# Build machine path to MS compiler
# Optional developer machine path to MS compiler
MS_IDE="C:\Program Files (x86)/Microsoft Visual Studio 9.0/Common7/IDE/devenv.exe"
#MS_IDE="C:/Program Files (x86)/Microsoft Visual Studio 9.0/Common7/IDE/devenv.exe"
#MS_IDE="C:/WINDOWS/Microsoft.NET/Framework/v3.5/MSBuild.exe"

ML = Addins/Excel/xquery
MS = Addins/Excel/Microsoft
MSS = MarkLogic_ExcelAddin
JS = Addins/Excel/javascript
XQY = Addins/Excel/xquery
CF = Addins/Excel/config
SAMPLES = Addins/Excel/Samples
CPF = Addins/Excel/cpf
DOCS = Addins/Excel/docs
JSDOCS = $(DOCS)/jsdocs
XQAPIDOCS = $(DOCS)/xquery-apidoc
INSTALL=Addins/Excel/install

BUILDS = builds
PUB_BUILD = $(ZIP_PREFIX)-$(SUFFIX)
ZIP_FILE = $(ZIP_PREFIX)-$(SUFFIX)

BUILD_DOCS = $(PUB_BUILD)/docs
BUILD_DOCS_JSDOC = $(BUILD_DOCS)/jsdocs
BUILD_DOCS_XQAPIDOC = $(BUILD_DOCS)/xquery-apidoc
BUILD_JS = $(PUB_BUILD)/js
BUILD_XQY =  $(PUB_BUILD)/xquery
BUILD_INSTALL = $(PUB_BUILD)/install
BUILD_CPF = $(PUB_BUILD)/cpf

BUILD_SAMPLES = $(PUB_BUILD)/Samples
BUILD_SAMPLES_JS = $(PUB_BUILD)/Samples/js
BUILD_SAMPLES_CSS = $(PUB_BUILD)/Samples/css
BUILD_SAMPLES_METADATA = $(PUB_BUILD)/Samples/metadata
BUILD_SAMPLES_SEARCH = $(PUB_BUILD)/Samples/search
BUILD_SAMPLES_SAVE = $(PUB_BUILD)/Samples/save

MS_PUB_BUILD = $(PUB_BUILD)/addin.deploy
MS_ROOT = $(MS)/MarkLogic_ExcelAddin
MS_MAIN_REL = $(MS)/MarkLogic_ExcelAddin/MarkLogic_ExcelAddin/bin/Release
MS_MLC_DIR = $(MS_ROOT)/MarkLogic_ExcelAddin_Setup
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
	mkdir $(BUILD_INSTALL)
	mkdir $(BUILD_CPF)
	mkdir $(BUILD_DOCS_JSDOC)
	mkdir $(BUILD_DOCS_JSDOC)/MLA
	mkdir $(BUILD_DOCS_XQAPIDOC)
	mkdir $(BUILD_DOCS_XQAPIDOC)/images
	mkdir $(BUILD_SAMPLES)
	mkdir $(BUILD_SAMPLES_JS)
	mkdir $(BUILD_SAMPLES_CSS)
	mkdir $(BUILD_SAMPLES_METADATA)
	mkdir $(BUILD_SAMPLES_SEARCH)
	mkdir $(BUILD_SAMPLES_SAVE)
	mkdir $(PUB_BUILD)/config
	mkdir $(MS_PUB_BUILD)
	cp $(CF)/*.idt $(PUB_BUILD)/config/.
	cp Addins/Excel/README.txt $(PUB_BUILD)
	cp LICENSE.txt $(PUB_BUILD)
	cp Addins/Excel/NOTICE.txt $(PUB_BUILD)
	#begin 323
	$(MS_IDE) $(MS_MLC_DIR)/MarkLogic_ExcelAddin_Setup.vdproj /build "Release"
	cp -r   $(MS_BUILD)/* $(MS_PUB_BUILD)/.
	#./setVersion patch $(JS)/MarkLogicExcelAddin.js $(BUILD_JS)/MarkLogicExcelAddin.js
	cp $(JS)/MarkLogicExcelAddin.js $(BUILD_JS)/MarkLogicExcelAddin.js
	cp $(JS)/MarkLogicExcelEventHandlers.js $(BUILD_JS)/MarkLogicExcelEventHandlers.js
	cp $(JS)/MarkLogicExcelEventSupport.js $(BUILD_JS)/MarkLogicExcelEventSupport.js
	cp $(JS)/MarkLogicExcelAddin.js  $(BUILD_SAMPLES_JS)/MarkLogicExcelAddin.js
	#./setVersion patch $(JS)/MarkLogicExcelAddin.js $(SAMPLES_JS)/MarkLogicExcelAddin.js
	#./setVersion patch $(ML)/word-processing-ml.xqy $(PUB_BUILD)/word-processing-ml.xqy
	#./setVersion patch $(ML)/package.xqy $(PUB_BUILD)/package.xqy
	#cp -r $(SAMPLES)/* $(BUILD_SAMPLES) 
	cp $(XQY)/spreadsheet-ml-support.xqy $(BUILD_XQY)
	cp $(INSTALL)/install.xqy $(BUILD_INSTALL)
	cp $(CPF)/*.xqy $(BUILD_CPF)
	cp $(CPF)/*.xml $(BUILD_CPF)
	cp $(SAMPLES)/default.xqy $(BUILD_SAMPLES)
	#cp $(SAMPLES)/README.txt $(BUILD_SAMPLES)
	#cp $(SAMPLES)/js/*.js $(BUILD_SAMPLES_JS) 
	cp $(SAMPLES)/css/*.css $(BUILD_SAMPLES_CSS) 
	cp $(SAMPLES)/metadata/*.js $(BUILD_SAMPLES_METADATA) 
	cp $(SAMPLES)/metadata/*.xqy $(BUILD_SAMPLES_METADATA) 
	cp $(SAMPLES)/metadata/*.css $(BUILD_SAMPLES_METADATA) 
	cp $(SAMPLES)/metadata/*.png $(BUILD_SAMPLES_METADATA) 
	cp $(SAMPLES)/search/*.js $(BUILD_SAMPLES_SEARCH) 
	cp $(SAMPLES)/search/*.xqy $(BUILD_SAMPLES_SEARCH) 
	cp $(SAMPLES)/search/*.gif $(BUILD_SAMPLES_SEARCH) 
	cp $(SAMPLES)/search/*.bmp $(BUILD_SAMPLES_SEARCH) 
	cp $(SAMPLES)/search/*.jpg $(BUILD_SAMPLES_SEARCH) 
	cp $(SAMPLES)/save/*.xqy $(BUILD_SAMPLES_SAVE) 
	cp $(SAMPLES)/save/*.js $(BUILD_SAMPLES_SAVE) 
	cp $(SAMPLES)/save/*.png $(BUILD_SAMPLES_SAVE) 
	cp $(JSDOCS)/*.css $(BUILD_DOCS_JSDOC)
	cp $(JSDOCS)/*.html $(BUILD_DOCS_JSDOC)
	cp $(JSDOCS)/MLA/*.html $(BUILD_DOCS_JSDOC)/MLA
	cp $(XQAPIDOCS)/*.html $(BUILD_DOCS_XQAPIDOC)
	cp $(XQAPIDOCS)/images/*.gif $(BUILD_DOCS_XQAPIDOC)/images
	cp $(XQAPIDOCS)/images/*.css $(BUILD_DOCS_XQAPIDOC)/images
	cp $(DOCS)/ToolkitForExcelGuide.docx $(BUILD_DOCS)
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

