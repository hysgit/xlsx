#!/usr/bin/env bash
mvn assembly:assembly
cp target/xlsx-1.0-SNAPSHOT-jar-with-dependencies.jar ~/tmp/change/txt.jar
cd ~/tmp/change/
java -jar txt.jar
cd -

