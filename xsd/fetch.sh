#!/bin/bash
# ECMA-376 transitional SpreadsheetML XSDs, taken from the Apache POI ooxml-schemas jar (Apache-2.0).
cd "$(dirname "$0")" && [ -f ooxml-schemas-1.4.jar ] || curl -sfLO https://repo1.maven.org/maven2/org/apache/poi/ooxml-schemas/1.4/ooxml-schemas-1.4.jar
rm -rf x && mkdir x && unzip -q -o ooxml-schemas-1.4.jar 'schemaorg_apache_xmlbeans/src/*.xsd' -d x
