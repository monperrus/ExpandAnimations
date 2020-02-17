#!/bin/bash
set -e
CURRENT_VERSION=`xmlstarlet sel -N oo="http://openoffice.org/extensions/description/2006" -t -v "//oo:version/@value" extension/description.xml`
MAJOR_VERSION=${CURRENT_VERSION%.*}
MINOR_VERSION=${CURRENT_VERSION#*.}
NEW_VERSION=$MAJOR_VERSION'.'$(($MINOR_VERSION+1))
echo "Current major version: "$MAJOR_VERSION
echo "Current minor version: "$MINOR_VERSION
echo "New version: "$NEW_VERSION
xmlstarlet ed --inplace -N oo="http://openoffice.org/extensions/description/2006" -u "//oo:version/@value" -v $NEW_VERSION extension/description.xml