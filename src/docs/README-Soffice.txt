StarOffice/OpenOffice Argouml plugin version 0.1.0
===================================================
Date: 2004-01-25
Author: Matti Pehrs

Installation
===================

The StarOffice/OpenOffice plugin is distributed as a zip 
file which you can install into a argouml home directory:

cd /opt
mkdir argouml
cd argouml;
tar xvfz /tmp/ArgoUML-0.14.1.tar.gz
unzip /tmp/argosoffice-0.1.0-20040125.zip
chmod 755 argouml.sh
./argouml.sh

Configuration
===================

To start ArgoUML correctly with all paths set I have 
provided a startup script: argouml.sh 
This script needs two environment variables set:

JAVA_HOME - Pointing to a Java SDK or JRE

OFFICE_HOME - Pointing to the StarOffice/OpenOffice root 

In order to make the StarOffice/OpenOffice plugin work you need to 
hade StarOffice/OpenOffice up and running and configured right:

-----------From OpenOffice.org1.1_SDK------------
Make the office listen

Java uses a TCP/IP socket to talk to the office. For Java clients, 
OpenOffice must be told to listen for TCP/IP connections using a 
special connection url parameter. There are two ways to achieve this, 
you can make the office listen always or just once.

To make the office listen whenever it is started, open the file 
<OfficePath>/share/registry/data/org/openoffice/Setup.xcu in an editor, 
and look for the element:
      <node oor:name="Office"/>
This element contains <prop/> elements. Insert the following <prop/> 
element on the same level as the existing elements:
      <prop oor:name="ooSetupConnectionURL" oor:type="xs:string">
       <value>socket,host=localhost,port=8100;urp;</value>
      </prop>

This setting configures OpenOffice.org to provide a socket on port 8100, 
where it will serve connections through the UNO remote protocol (urp). 
If port 8100 is already in use on your machine, it may be necessary to 
adjust the port number. Block port 8100 for connections from outside your 
network in your firewall. If you have a OpenOffice.org network installation, 
this setting will affect all users. To make only a particular user installation 
listen for connections create a file <OfficePath>/user/registry/data/org/openoffice/Setup.xcu 
with the same structure as the file above and add the element 
<prop oor:name=" ooSetupConnectionURL"/> as shown above.

-----------From OpenOffice.org1.1_SDK------------

Limitations
======================
- The plugin assumes the port 8100 for now. This will be configurable in 
  future versions.

- The Star/Open-Office Plugin can generate all the different types of
  digrams supported by ArgoUML. Rounded Rects are not supporte yet so
  State diagrams will not look correct. Fill color support is not fully 
  implemented which will generate strange colors on some figs.
  The Dictionary is still only generating a dictionary for Use Cases/Actors 
  and Deployment diagram Components.

Links
======================

ArgoUML - http://argouml.tigris.org

StarOffice - http://wwws.sun.com/software/star/staroffice/index.html

OpenOffice - http://www.openoffice.org/

OpenOffice/StarOffice SDK - http://www.openoffice.org/dev_docs/source/sdk/index.html


License
======================
// Copyright (c) 2003, Matti Pehrs
// All rights reserved.
//
// Redistribution and use in source and binary forms, with or without modification, 
// are permitted provided that the following conditions are met:
//
//     * Redistributions of source code must retain the above copyright notice, 
//       this list of conditions and the following disclaimer.
//     * Redistributions in binary form must reproduce the above copyright notice, 
//       this list of conditions and the following disclaimer in the documentation 
//       and/or other materials provided with the distribution.
//     * Neither the name of the Matti Pehrs nor the names of its contributors may 
//       be used to endorse or promote products derived from this software without 
//       specific prior written permission.
//
// THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND 
// ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED 
// WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. 
// IN NO EVENT SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, 
// INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, 
// BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, 
// DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY 
// OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE 
// OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED 
// OF THE POSSIBILITY OF SUCH DAMAGE.

