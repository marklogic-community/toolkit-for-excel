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

ML = Addins/Word/xquery
MS = Addins/Word/Microsoft
MSS = MarkLogic_WordAddin
JS = Addins/Word/javascript
CF = Addins/Word/config
BUILDS = builds
PUB_BUILD = $(BUILDS)/Word
MS_PUB_BUILD = $(PUB_BUILD)/addin.deploy
MS_ROOT = $(MS)/MarkLogic_WordAddin
MS_MLC_DIR = $(MS_ROOT)/MarkLogic_WordAddin/bin
MS_BUILD = $(MS_MLC_DIR)/Debug/app.publish
TEMP = temp
#
# Microsoft build
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
	mkdir $(PUB_BUILD)/config
	mkdir $(MS_PUB_BUILD)
	cp $(CF)/*.reg $(PUB_BUILD)/config/.
	cp README.txt $(PUB_BUILD)
	cp  $(MS_ROOT)/$(MSS)/UserControl1.cs  $(TEMP)/UserControl1.cs.bak
	./setVersion patch $(MS_ROOT)/$(MSS)/UserControl1.cs  $(MS_ROOT)/$(MSS)/UserControl2.cs
	rm $(MS_ROOT)/$(MSS)/UserControl1.cs
	mv $(MS_ROOT)/$(MSS)/UserControl2.cs $(MS_ROOT)/$(MSS)/UserControl1.cs
	build-addin.bat
	mv $(TEMP)/UserControl1.cs.bak  $(MS_ROOT)/$(MSS)/UserControl1.cs
	cp -r   $(MS_BUILD)/* $(MS_PUB_BUILD)/.
	./setVersion patch $(JS)/MarkLogicWordAddin.js $(PUB_BUILD)/MarkLogicWordAddin.js
	./setVersion patch $(ML)/word-processing-ml.xqy $(PUB_BUILD)/word-processing-ml.xqy
	./setVersion patch $(ML)/package.xqy $(PUB_BUILD)/package.xqy
	@echo Create zip file $(ZIP_PREFIX)_$(SUFFIX).zip
	(cd builds; zip -r ../$(ZIP_PREFIX).zip Word/*)
	mv $(ZIP_PREFIX).zip $(ZIP_PREFIX)-$(SUFFIX).zip

clean:
	  rm -rf $(BUILDS)
	  rm -rf $(MS_BUILD)
	  rm -rf $(TEMP)
	  rm ./*.zip

