#!/bin/bash
set -e
CURRENT_VERSION=`sed -n 's/.*<version value="\([^"]*\)".*/\1/p' extension/description.xml`
MAJOR_VERSION=${CURRENT_VERSION%.*}
MINOR_VERSION=${CURRENT_VERSION#*.}
NEW_VERSION=$MAJOR_VERSION'.'$(($MINOR_VERSION+1))
echo "Current major version: "$MAJOR_VERSION
echo "Current minor version: "$MINOR_VERSION
echo "New version: "$NEW_VERSION
sed -i "s/<version value=\"$CURRENT_VERSION\"\/>/<version value=\"$NEW_VERSION\"\/>/" extension/description.xml
