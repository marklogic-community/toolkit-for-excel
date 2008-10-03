#
# Makefile
#
# Instructions:
#  
#

MAJ_VER := `cat MAJOR_VERSION`
MIN_VER := `cat MINOR_VERSION`
DATE := `date +%Y%m%d`
SUFFIX := $(MAJ_VER).$(MIN_VER)-$(DATE)
ZIP_PREFIX = MarkLogic_WordAddin

# Build machine path to MS compiler
# Optional developer machine path to MS compiler
#MS_IDE="C:/Program Files/Microsoft Visual Studio 9.0/Common7/IDE/devenv.exe"
MS_IDE="C:/Program Files (x86)/Microsoft Visual Studio 9.0/Common7/IDE/devenv.exe"
#MS_IDE="C:/WINDOWS/Microsoft.NET/Framework/v3.5/MSBuild.exe"

ML = Addins/Word/xquery
MS = Addins/Word/Microsoft
MSS = MarkLogic_WordAddin
JS = Addins/Word/javascript
CF = Addins/Word/config
SAMPLES = Addins/Word/Samples
SAMPLES_JS = $(SAMPLES)/js

BUILDS = builds
PUB_BUILD = $(BUILDS)/Word

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
	./setVersion patch $(JS)/MarkLogicWordAddin.js $(PUB_BUILD)/MarkLogicWordAddin.js
	./setVersion patch $(JS)/MarkLogicWordAddin.js $(SAMPLES_JS)/MarkLogicWordAddin.js
	./setVersion patch $(ML)/word-processing-ml.xqy $(PUB_BUILD)/word-processing-ml.xqy
	./setVersion patch $(ML)/package.xqy $(PUB_BUILD)/package.xqy
	#cp -r $(SAMPLES)/* $(BUILD_SAMPLES) 
	cp $(SAMPLES)/default.xqy $(BUILD_SAMPLES)
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
	#cp -r $(SAMPLES)/modules/*.js $(BUILD_SAMPLES_JS) 
	@echo Create zip file $(ZIP_PREFIX)_$(SUFFIX).zip
	(cd builds; zip -r ../$(ZIP_PREFIX).zip Word/*)
	mv $(ZIP_PREFIX).zip $(ZIP_PREFIX)-$(SUFFIX).zip

clean:
	  rm -rf $(BUILDS)
	  rm -rf $(MS_BUILD)/*
	  rm -rf $(MS_MAIN_REL)/*
	  rm -rf $(TEMP)
	  rm ./*.zip

