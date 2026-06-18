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
	cd extension; find . -exec touch -t 200001010000 {} +; zip -X -r ../$@ .

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
	test -f test/test-ExpandAnimations-expanded.odp
	@if unzip -p test/test-ExpandAnimations-expanded.odp content.xml | grep -E '<anim:|presentation:node-type'; then \
		echo "Expanded ODP still contains animations"; \
		exit 1; \
	fi
	@if unzip -p test/test-ExpandAnimations-expanded.odp content.xml | grep -E 'text:page-count|text:page-number'; then \
		echo "Expanded ODP still contains dynamic page fields"; \
		exit 1; \
	fi
	bash pdf-diff.sh
	pdffonts test/test-ExpandAnimations.pdf | awk 'NR > 2 && $$(NF-4) != "yes" { print; bad=1 } END { exit bad }'
	EXPANDANIMATIONS_INPUT=$(CURDIR)/test/BadVerticalAlign.odp libreoffice -env:UserInstallation=$(TEST_USER_INSTALLATION) --headless macro:///ExpandAnimations.ExpandAnimations.CommandLine
	test -f test/BadVerticalAlign-expanded.odp
	unzip -p test/BadVerticalAlign-expanded.odp content.xml | grep -q '<text:s/>'
	pdffonts test/BadVerticalAlign.pdf | awk 'NR > 2 && $$(NF-4) != "yes" { print; bad=1 } END { exit bad }'
	EXPANDANIMATIONS_INPUT=$(CURDIR)/test/UnsupportedEmphasis.odp libreoffice -env:UserInstallation=$(TEST_USER_INSTALLATION) --headless macro:///ExpandAnimations.ExpandAnimations.CommandLine
	test -f test/UnsupportedEmphasis-expanded.odp
	test -f test/UnsupportedEmphasis.pdf
	pdfinfo test/UnsupportedEmphasis.pdf | grep -Eq '^Pages:[[:space:]]+7$$'
	pdffonts test/UnsupportedEmphasis.pdf | awk 'NR > 2 && $$(NF-4) != "yes" { print; bad=1 } END { exit bad }'
	@if unzip -p test/UnsupportedEmphasis-expanded.odp content.xml | grep -E '<anim:|presentation:node-type'; then \
		echo "Expanded unsupported-emphasis ODP still contains animations"; \
		exit 1; \
	fi
	EXPANDANIMATIONS_INPUT=$(CURDIR)/test/links.odp libreoffice -env:UserInstallation=$(TEST_USER_INSTALLATION) --headless macro:///ExpandAnimations.ExpandAnimations.CommandLine
	test -f test/links-expanded.odp
	test -f test/links.pdf
	pdfinfo test/links.pdf | grep -Eq '^Pages:[[:space:]]+4$$'
	pdffonts test/links.pdf | awk 'NR > 2 && $$(NF-4) != "yes" { print; bad=1 } END { exit bad }'
	unzip -p test/links-expanded.odp content.xml | grep -q 'xlink:href="#Slide: 3"'
	@if unzip -p test/links-expanded.odp content.xml | grep -q 'presentation:visibility="hidden"'; then \
		echo "Expanded ODP still contains hidden slides"; \
		exit 1; \
	fi
	@if grep -aE '/URI\(#|/GoToR' test/links.pdf; then \
		echo "Expanded PDF still contains externalized internal links"; \
		exit 1; \
	fi
	grep -a '/Dest\[' test/links.pdf
	grep -a '/Dest\[7 0 R/FitR' test/links.pdf
	libreoffice -env:UserInstallation=$(TEST_USER_INSTALLATION) --headless --convert-to pdf --outdir test test/links-expanded.odp
	test -f test/links-expanded.pdf
	@if grep -aE '/URI\(#|/GoToR' test/links-expanded.pdf; then \
		echo "Re-exported expanded PDF still contains externalized internal links"; \
		exit 1; \
	fi
	grep -a '/Dest\[17 0 R/FitR' test/links-expanded.pdf

clean:
	rm -f extension/ExpandAnimations/ExpandAnimations.xba
	rm -f test/test-ExpandAnimations.pdf
	rm -f test/BadVerticalAlign.pdf
	rm -f test/links.pdf
	rm -f test/UnsupportedEmphasis.pdf
	rm -f test/*-expanded.pdf
	rm -f test/*-expanded.odp
	rm -f test/*.txt
	rm -f test/.~lock*
	rm -f dist/$(EXTENSIONNAME)-$(VERSION).oxt
	rm -rf $(TEST_PROFILE)
