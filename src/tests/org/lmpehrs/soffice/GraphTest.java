//
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
//
package org.lmpehrs.soffice;

// base classes
import java.io.File;

import com.sun.star.drawing.XShape;

/**
 * Test Class to generate Graphics in Star/OpenOffice
 *
 * @author  Matti Pehrs
 * @version $Id$
 */
public class GraphTest {
    
    public GraphTest() {
    }
    
    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
	System.out.println("GraphTest...");
        GraphTest test = new GraphTest();
        try {
            test.drawSDrawDoc();
        } catch (java.lang.Exception e){
            e.printStackTrace();
        } finally {
            System.exit(0);
        }
    }

    public void drawSDrawDoc() throws java.lang.Exception {

	java.io.File templateFile = 
	    new java.io.File("templates"+File.separator+"argouml.std");
	StringBuffer sTmp = new StringBuffer("file:///");
	sTmp.append(templateFile.getCanonicalPath().replace('\\', '/'));
	String templateDoc = sTmp.toString();
	
	SDrawDoc doc = new SDrawDoc();
	doc.connect(templateDoc);

	doc.setVirtualWidth(1640);
	doc.setVirtualHeight(1450);

	XShape uc1 = doc.drawUseCase(104, 256, 
				     150, 100, 
				     "UseCase 1");
	
	XShape actor1 = doc.drawActor(600, 200, 
				      50, 150,
				      "Actor 1");
	
	XShape ass1 = doc.drawAssociation(actor1, uc1, false, true,
					   "Association", "From Role", "To Role");


	XShape classA = doc.drawClass(500, 500, 
				      200, 150,
				      "ClassA");

	XShape classAInheritance = doc.addInheritanceTriangle(classA);
	XShape classB = doc.drawClass(300, 800, 
				      200, 150,
				      "ClassB");
	XShape classC = doc.drawClass(600, 800, 
				      200, 150,
				      "ClassC");
	XShape inheritacneLine1 = doc.drawRectLine(classB, classAInheritance);
	XShape inheritacneLine2 = doc.drawRectLine(classC, classAInheritance);


	XShape comp1 = doc.drawComponent(1000, 800, 
					 200, 150,
					 "Component 1");
	
	
	XShape comp2 = doc.drawComponent(1200, 500, 
					 200, 150,
					 "Component 2");

	XShape compAss = doc.drawAssociation(comp1, comp2, false, true,
					      null, null, null);
	Soffice.setDashedLine(compAss);
    }
    
}
