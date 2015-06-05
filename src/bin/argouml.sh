#!/bin/bash
#
# Shell script to launch ArgoUML on Unix systems.  Mostly borrowed from
# the Apache Ant project.
#
# The idea is that you can put a softlink in your "bin" directory back
# to this file in the ArgoUML install directory and this script will
# use the link to find the jars that it needs, e.g.:
#
# ln -s /usr/local/ArgoUML/argouml.sh /usr/local/bin/argo
#
# 2002-02-25 toby@caboteria.org

## resolve links - $0 may be a link to ArgoUML's home
PRG=$0
progname=`basename $0`

if [ -z $JAVA_HOME ] ; then
      JAVA_HOME=/usr/java/j2sdk1.4.2
      echo "Warning: JAVA_HOME not set! using ${JAVA_HOME}"
fi

if [ -z $OFFICE_HOME ] ; then
      OFFICE_HOME=/usr/local/OpenOffice.org1.1.0
      echo "Warning: OFFICE_HOME not set! using ${OFFICE_HOME}"
fi

while [ -h "$PRG" ] ; do
  ls=`ls -ld "$PRG"`
  link=`expr "$ls" : '.*-> \(.*\)$'`
  if expr "$link" : '.*/.*' > /dev/null; then
      PRG="$link"
  else
      PRG="`dirname $PRG`/$link"
  fi
done

# ARGO_HOME
ARGO_HOME=`dirname $PRG`
CP=$ARGO_HOME

#Core JARS
JARS=`ls $ARGO_HOME/*.jar`
for f in $JARS; do 
    CP=$CP":"$f
done

# Extensions JARS
LIBS=`ls $ARGO_HOME/lib`
for f in $LIBS; do 
    CP=$CP":"$ARGO_HOME/lib/$f
done

# SOFFICE_JARs
CP=$CP":"$OFFICE_HOME/program/classes/juh.jar":"$OFFICE_HOME/program/classes/jurt.jar":"$OFFICE_HOME/program/classes/ridl.jar":"$OFFICE_HOME/program/classes/sandbox.jar":"$OFFICE_HOME/program/classes/unoil.jar

#java -jar `dirname $PRG`/argouml.jar "$@"
java -Xmx500m -classpath $CP org.argouml.application.Main "$@"

