#!/bin/sh
# ./gradlew clean
./gradlew build
java -jar app/build/libs/app-all.jar files/ $1
