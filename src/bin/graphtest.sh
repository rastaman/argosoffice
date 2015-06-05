#!/bin/sh
BINDIR=`dirname $0`/..

SOFFICE_HOME=/ext/staroffice7
OFFICE_CLASSES_DIR=$SOFFICE_HOME/program/classes

CP=build/classes:$OFFICE_CLASSES_DIR/jurt.jar:$OFFICE_CLASSES_DIR/unoil.jar:$OFFICE_CLASSES_DIR/ridl.jar:$OFFICE_CLASSES_DIR/sandbox.jar:$OFFICE_CLASSES_DIR/juh.jar

$JAVA_HOME/bin/java -cp $CP org.lmpehrs.soffice.GraphTest $@
