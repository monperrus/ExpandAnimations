# Makefile for building LibreOffice/OpenOffice extensions

EXTENSIONNAME=ExpandAnimations
VERSION=0.1

all: dist/$(EXTENSIONNAME)-$(VERSION).oxt

dist/$(EXTENSIONNAME)-$(VERSION).oxt: $(shell find extension)
	cd extension; zip -r ../$@ .

extension/ExpandAnimations/ExpandAnimations.xba: src/ExpandAnimations.bas
	echo '<?xml version="1.0" encoding="UTF-8"?>' > $@
	echo '<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">' >> $@
	echo '<script:module xmlns:script="http://openoffice.org/2000/script" script:name="ExpandAnimations" script:language="StarBasic">' >> $@
	perl -MHTML::Entities -ne 'print encode_entities($$_)' $^ >> $@
	echo '</script:module>' >> $@

