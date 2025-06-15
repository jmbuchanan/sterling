#!/bin/bash

jpackage --input target/ \
  --name sterling \
  --main-jar sterling-1.0-SNAPSHOT.jar \
  --main-class com.sterling.automation.App \
  --type deb \
  --java-options '--enable-preview'
