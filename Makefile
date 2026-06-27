# Makefile for building LibreOffice/OpenOffice extensions

EXTENSIONNAME=ExpandAnimations
VERSION=$(shell sed -n 's/.*<version value="\([^"]*\)".*/\1/p' extension/description.xml)
TEST_PROFILE=$(CURDIR)/.test-libreoffice-profile
TEST_USER_INSTALLATION=file://$(TEST_PROFILE)
ifeq ($(VERSION),)
$(error Could not read extension version from extension/description.xml)
endif

all: dist/$(EXTENSIONNAME)-$(VERSION).oxt

validate: all
	unzip -t dist/$(EXTENSIONNAME)-$(VERSION).oxt
	unzip -Z1 dist/$(EXTENSIONNAME)-$(VERSION).oxt | grep -Fxq META-INF/manifest.xml
	unzip -Z1 dist/$(EXTENSIONNAME)-$(VERSION).oxt | grep -Fxq description.xml
	unzip -Z1 dist/$(EXTENSIONNAME)-$(VERSION).oxt | grep -Fxq Addons.xcu
	unzip -Z1 dist/$(EXTENSIONNAME)-$(VERSION).oxt | grep -Fxq ExpandAnimations/ExpandAnimations.xba

compatibility:
	libreoffice --version
	unopkg --version

dist/$(EXTENSIONNAME)-$(VERSION).oxt: extension/ExpandAnimations/ExpandAnimations.xba
	mkdir -p dist
	cd extension; zip -r ../$@ .

extension/ExpandAnimations/ExpandAnimations.xba: src/ExpandAnimations.bas
	echo '<?xml version="1.0" encoding="UTF-8"?>' > $@
	echo '<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">' >> $@
	echo '<script:module xmlns:script="http://openoffice.org/2000/script" script:name="ExpandAnimations" script:language="StarBasic">' >> $@
	perl -MHTML::Entities -ne 'print encode_entities($$_)' $^ >> $@
	echo '</script:module>' >> $@

install: all
	unopkg add -s dist/$(EXTENSIONNAME)-$(VERSION).oxt

uninstall:
	-unopkg remove vnd.basicaddonbuilder.expandanimations

test: all
	rm -rf $(TEST_PROFILE)
	mkdir -p $(TEST_PROFILE)
	unopkg -env:UserInstallation=$(TEST_USER_INSTALLATION) add -f -s dist/$(EXTENSIONNAME)-$(VERSION).oxt
	libreoffice -env:UserInstallation=$(TEST_USER_INSTALLATION) --headless test/test-ExpandAnimations.odp macro:///ExpandAnimations.ExpandAnimations.test
	bash pdf-diff.sh

clean:
	rm -f extension/ExpandAnimations/ExpandAnimations.xba
	rm -f test/test-ExpandAnimations.pdf
	rm -f test/*.txt
	rm -f test/.~lock*
	rm -f dist/$(EXTENSIONNAME)-$(VERSION).oxt
	rm -rf $(TEST_PROFILE)
