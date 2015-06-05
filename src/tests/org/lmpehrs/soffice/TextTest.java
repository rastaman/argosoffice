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

import com.sun.star.awt.Point;
import com.sun.star.awt.Size;
import com.sun.star.beans.XPropertySet;
import com.sun.star.beans.XPropertySetInfo;
import com.sun.star.drawing.TextFitToSizeType;
import com.sun.star.drawing.XDrawPageSupplier;
import com.sun.star.drawing.XShape;
import com.sun.star.drawing.XShapes;
import com.sun.star.text.ControlCharacter;
import com.sun.star.text.SizeType;
import com.sun.star.text.TextContentAnchorType;
import com.sun.star.text.WrapTextMode;
import com.sun.star.text.XText;
import com.sun.star.text.XTextContent;
import com.sun.star.text.XTextDocument;
import com.sun.star.text.XTextFrame;
import com.sun.star.uno.UnoRuntime;

/**
 * Test Class to generate Graphics in Star/OpenOffice
 *
 * @author  Matti Pehrs
 * @version $Id$
 */
public class TextTest {
    
    public TextTest() {
    }
    
    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
	System.out.println("TextTest...");
        TextTest test = new TextTest();
        try {
            test.draw();
        } catch (java.lang.Exception e){
            e.printStackTrace();
        } finally {
            System.exit(0);
        }
    }

    public void draw() throws java.lang.Exception {

	SWriterDoc doc = new SWriterDoc();

	// Load template
	java.io.File templateFile = 
	    new java.io.File("templates"+File.separator+"argouml.stw");
	StringBuffer sTmp = new StringBuffer("file:///");
	sTmp.append(templateFile.getCanonicalPath().replace('\\', '/'));
	String templateDoc = sTmp.toString();

	doc.connect(templateDoc);

	if(1 == 0) {
	    doc.addParagraph("Test Document", "Heading 1");
	    doc.addParagraph("Document Generated "+(new java.util.Date()).toString());
	    
	    
	    // Insert Frame
	    java.lang.Object frameObject = 
		doc.getDocumentFactory().createInstance ("com.sun.star.text.TextFrame");
	    System.out.println("frameObject="+frameObject);
	    XTextFrame xFrame =
		(XTextFrame) UnoRuntime.queryInterface (XTextFrame.class, 
							frameObject);
	    XText tfText = xFrame.getText();
	    System.out.println("xFrame="+xFrame);
	    XShape frameShape = (XShape) UnoRuntime.queryInterface(XShape.class, xFrame);
	// Access the XPropertySet interface of the TextFrame
	    XPropertySet xFrameProps = 
		(XPropertySet)UnoRuntime.queryInterface(XPropertySet.class, xFrame );
	    frameShape.setSize(new Size(25400, 11400));
	    // Set the AnchorType to com.sun.star.text.TextContentAnchorType.AS_CHARACTER
	    xFrameProps.setPropertyValue( "AnchorType", TextContentAnchorType.AT_PARAGRAPH );
	    xFrameProps.setPropertyValue("ZOrder", new Integer(500));
	    xFrameProps.setPropertyValue("FrameIsAutomaticHeight", new Boolean(false));
	    xFrameProps.setPropertyValue("SizeType", new Short(SizeType.FIX));
	    // Go to the end of the text document
	    doc.getXTextCursor().gotoEnd(false);
	    // Insert a new paragraph
	    doc.getXText().insertControlCharacter(doc.getXTextCursor(), ControlCharacter.PARAGRAPH_BREAK, false );
	    // Then insert the new frame
	    doc.getXText().insertTextContent(doc.getXTextCursor(), xFrame, false);
	    
	    // Add some text
	    tfText.insertString(tfText.getEnd(), "...", false );
	    tfText.insertControlCharacter(tfText.getEnd(), ControlCharacter.PARAGRAPH_BREAK, false );
	    
	    // Create rectangle 1
	    Object rect1obj = doc.getDocumentFactory().createInstance("com.sun.star.drawing.RectangleShape");
	    XShape rect1 = (XShape)UnoRuntime.queryInterface(XShape.class, rect1obj);
	    rect1.setPosition( new Point( 1000, 1000 ) );
	    rect1.setSize(new Size(1000, 1000));
	    
	    // Create rectangle 2
	    Object rect2obj = doc.getDocumentFactory().createInstance("com.sun.star.drawing.RectangleShape");
	    XShape rect2 = (XShape)UnoRuntime.queryInterface(XShape.class, rect2obj);
	    rect2.setPosition( new Point( 2000, 2000 ) );
	    rect2.setSize(new Size(1000, 1000));
	    
	    //
	    // Create and add group
	    //
	    Object groupobj = doc.getDocumentFactory().createInstance("com.sun.star.drawing.GroupShape");
	    XShape group = (XShape)UnoRuntime.queryInterface(XShape.class, groupobj);
	    group.setPosition( new Point( 4000, 1000 ) );
	    // Add Group Shape to text doc
	    // query for the shape collection of xDrawPage
	    XTextContent groupTextContent = 
		(XTextContent)UnoRuntime.queryInterface(XTextContent.class, groupobj);
	    tfText.insertTextContent(tfText.getEnd(), groupTextContent, false);
	    // Set Properties
	    XPropertySet groupProps = (XPropertySet)UnoRuntime.queryInterface(XPropertySet.class, groupobj);
	    // wrap text inside shape
	    groupProps.setPropertyValue("AnchorType", TextContentAnchorType.AT_PARAGRAPH);
	    groupProps.setPropertyValue("ZOrder", new Integer(1000000));
	    XShapes groupShapes = (XShapes)UnoRuntime.queryInterface(XShapes.class, groupobj);
	    groupShapes.add(rect1);
	    groupShapes.add(rect2);
	    
	    // Set Rect 1 Properties
	    XPropertySet rect1Props = (XPropertySet)UnoRuntime.queryInterface(XPropertySet.class, rect1obj);
	    // wrap text inside shape
	    rect1Props.setPropertyValue("TextContourFrame", new Boolean(true));
	    // rect1Props.setPropertyValue("AnchorType", TextContentAnchorType.AT_PARAGRAPH);
	    // rect1Props.setPropertyValue("ZOrder", new Integer(0));
	    XText rect1Text = (XText)UnoRuntime.queryInterface(XText.class, rect1obj);
	    rect1Text.insertString(rect1Text.getEnd(), "Hello", false );
	    
	    
	    // Set Rect 2 Properties
	    XPropertySet rect2Props = (XPropertySet)UnoRuntime.queryInterface(XPropertySet.class, rect2obj);
	    // wrap text inside shape
	    rect2Props.setPropertyValue("TextContourFrame", new Boolean(true));
	    // rect2Props.setPropertyValue("AnchorType", TextContentAnchorType.AT_PARAGRAPH);
	    // rect2Props.setPropertyValue("ZOrder", new Integer(0));
	    XText rect2Text = (XText)UnoRuntime.queryInterface(XText.class, rect2obj);
	    rect2Text.insertString(rect2Text.getEnd(), "Rect 2", false );
	    
	    
	    // Text
	    doc.getXText().insertString(doc.getXTextCursor(), 
					"This is a small Paragraph before the graphics group", 
					false );
	    
	    
	    //
	    // Insert Group 
	    // 
	    // Create bg rect
	    Object bgrectobj = doc.getDocumentFactory().createInstance("com.sun.star.drawing.RectangleShape");
	    XShape bgrect = (XShape)UnoRuntime.queryInterface(XShape.class, bgrectobj);
	    bgrect.setPosition( new Point(0, 0) );
	    bgrect.setSize(new Size(doc.getFrameWidth(), 13000));
	    
	    // Create rectangle 3
	    Object rect3obj = doc.getDocumentFactory().createInstance("com.sun.star.drawing.RectangleShape");
	    XShape rect3 = (XShape)UnoRuntime.queryInterface(XShape.class, rect3obj);
	    rect3.setPosition( new Point( 3000, 3000 ) );
	    rect3.setSize(new Size(3000, 3000));
	    // Set Rect 3 Properties
	    
	    // Create rectangle 4
	    Object rect4obj = doc.getDocumentFactory().createInstance("com.sun.star.drawing.RectangleShape");
	    XShape rect4 = (XShape)UnoRuntime.queryInterface(XShape.class, rect4obj);
	    rect4.setPosition( new Point( 14000, 4000 ) );
	    rect4.setSize(new Size(1000, 1000));
	    
	    //
	    // Create and add group
	    //
	    Object mgroupobj = doc.getDocumentFactory().createInstance("com.sun.star.drawing.GroupShape");
	    XShape mgroup = (XShape)UnoRuntime.queryInterface(XShape.class, mgroupobj);
	    // mgroup.setPosition( new Point( 0, 0 ) );
	    // mgroup.setSize( new Size( 25000, 15000 ) );
	    // Add Group Shape to text doc
	    XPropertySet mgroupProps = (XPropertySet)UnoRuntime.queryInterface(XPropertySet.class, mgroupobj);
	    // wrap text inside shape
	    mgroupProps.setPropertyValue("AnchorType", TextContentAnchorType.AS_CHARACTER);
	    mgroupProps.setPropertyValue("TextWrap", WrapTextMode.NONE);
	    // query for the shape collection of xDrawPage
	    XTextContent mgroupTextContent = 
		(XTextContent)UnoRuntime.queryInterface(XTextContent.class, mgroupobj);
	    doc.getXText().insertTextContent(doc.getXText().getEnd(), mgroupTextContent, false);
	    // Set Properties
	    XShapes mgroupShapes = (XShapes)UnoRuntime.queryInterface(XShapes.class, mgroupobj);
	    mgroupShapes.add(bgrect);
	    mgroupShapes.add(rect3);
	    mgroupShapes.add(rect4);
	    
	    
	    // Set BG Properties
	    XPropertySet bgrectProps = (XPropertySet)UnoRuntime.queryInterface(XPropertySet.class, bgrectobj);
	    bgrectProps.setPropertyValue("FillColor", new Integer(0xFFFFFF));
	    bgrectProps.setPropertyValue("TextContourFrame", new Boolean(false));
	    XPropertySet rect3Props = (XPropertySet)UnoRuntime.queryInterface(XPropertySet.class, rect3obj);
	    // wrap text inside shape
	    rect3Props.setPropertyValue("TextContourFrame", new Boolean(true));
	    rect3Props.setPropertyValue("TextFitToSize", TextFitToSizeType.PROPORTIONAL );
	    rect3Props.setPropertyValue("FillColor", new Integer(0xFFFFFF));
	    // rect3Props.setPropertyValue("AnchorType", TextContentAnchorType.AT_PARAGRAPH);
	    // rect3Props.setPropertyValue("ZOrder", new Integer(0));
	    XText rect3Text = (XText)UnoRuntime.queryInterface(XText.class, rect3obj);
	    rect3Text.insertString(rect3Text.getEnd(), "Rect 3", false );
	    rect3.setPosition( new Point( 3000, 3000 ) );
	    rect3.setSize(new Size(3000, 3000));
	    // Set Rect 4 Properties
	    XPropertySet rect4Props = (XPropertySet)UnoRuntime.queryInterface(XPropertySet.class, rect4obj);
	    // wrap text inside shape
	    rect4Props.setPropertyValue("FillColor", new Integer(0xFFFFFF));
	    rect4Props.setPropertyValue("TextFitToSize", TextFitToSizeType.PROPORTIONAL );
	    rect4Props.setPropertyValue("TextContourFrame", new Boolean(true));
	    // rect4Props.setPropertyValue("AnchorType", TextContentAnchorType.AT_PARAGRAPH);
	    // rect4Props.setPropertyValue("ZOrder", new Integer(0));
	    XText rect4Text = (XText)UnoRuntime.queryInterface(XText.class, rect4obj);
	    rect4Text.insertString(rect4Text.getEnd(), "Rect 4", false );
	}

	//
	// FRAME 2: Does not work...
	//
	if(1 == 0) {
	    
	    // Insert Frame 2
	    java.lang.Object frame2Object = 
		doc.getDocumentFactory().createInstance ("com.sun.star.text.TextFrame");
	    System.out.println("frame2Object="+frame2Object);
	    XTextFrame xFrame2 =
		(XTextFrame) UnoRuntime.queryInterface (XTextFrame.class, 
							frame2Object);
	    System.out.println("xFrame2="+xFrame2);
	    XText tfText2 = xFrame2.getText();
	    // XTextCursor tfText2Cursor = tfText2.createTextCursor();
	    XShape frame2Shape = (XShape) UnoRuntime.queryInterface(XShape.class, xFrame2);
	    // Access the XPropertySet interface of the TextFrame
	    XPropertySet xFrame2Props = 
		(XPropertySet)UnoRuntime.queryInterface(XPropertySet.class, xFrame2 );
	    frame2Shape.setSize(new Size(25400, 11400));
	    // Set the AnchorType to com.sun.star.text.TextContentAnchorType.AS_CHARACTER
	    xFrame2Props.setPropertyValue( "AnchorType", TextContentAnchorType.AT_PARAGRAPH );
	    xFrame2Props.setPropertyValue("ZOrder", new Integer(500));
	    xFrame2Props.setPropertyValue("FrameIsAutomaticHeight", new Boolean(false));
	    xFrame2Props.setPropertyValue("SizeType", new Short(SizeType.FIX));
	    // Go to the end of the text document
	    doc.getXTextCursor().gotoEnd(false);
	    // Insert a new paragraph
	    doc.getXText().insertControlCharacter(doc.getXTextCursor(), ControlCharacter.PARAGRAPH_BREAK, false );
	    // Then insert the new frame
	    doc.getXText().insertTextContent(doc.getXTextCursor(), xFrame2, false);

	    // Add some text
	    tfText2.insertString(tfText2.getEnd(), "...", false );
	    tfText2.insertControlCharacter(tfText2.getEnd(), ControlCharacter.PARAGRAPH_BREAK, false );
	    
	    
	    //  	    XParagraphCursor xParaCursor = (XParagraphCursor)
	    //  		UnoRuntime.queryInterface(XParagraphCursor.class, 
	    //  					  tfText2.getEnd());
 	    // Position the cursor in the 2nd paragraph
 	    // xParaCursor.gotoPreviousParagraph(false);
 	    // Create a RectangleShape using the document factory
 	    XShape xRect = 
 		(XShape) UnoRuntime.queryInterface(XShape.class, 
 						   doc.getDocumentFactory().createInstance("com.sun.star.drawing.RectangleShape"));
 	    // Create an EllipseShape using the document factory
 	    XShape xEllipse = 
 		(XShape) UnoRuntime.queryInterface(XShape.class, 
 						   doc.getDocumentFactory().createInstance("com.sun.star.drawing.EllipseShape"));
 	    // Set the size of both the ellipse and the rectangle
 	    xRect.setSize(new Size(4000, 10000));
 	    xEllipse.setSize(new Size(3000, 6000));
 	    // Set the position of the Rectangle to the right of the ellipse
 	    xRect.setPosition(new Point(6100, 0));
 	    // Get the XPropertySet interfaces of both shapes
 	    XPropertySet xRectProps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xRect);
 	    XPropertySet xEllipseProps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xEllipse);
 	    // And set the AnchorTypes of both shapes to 'AT_PARAGRAPH'
 	    xRectProps.setPropertyValue("AnchorType", TextContentAnchorType.AT_PARAGRAPH);
 	    xEllipseProps.setPropertyValue("AnchorType", TextContentAnchorType.AT_PARAGRAPH);
 	    // Access the XDrawPageSupplier interface of the document
	    XTextDocument frameDoc = (XTextDocument)UnoRuntime.queryInterface(XTextDocument.class,
									      frame2Object);
 	    XDrawPageSupplier xDrawPageSupplier = 
 		(XDrawPageSupplier) UnoRuntime.queryInterface(XDrawPageSupplier.class, 
 							      frameDoc); // doc.getXTextDocument());
 	    // Get the XShapes interface of the draw page
 	    XShapes xShapes = 
 		(XShapes) UnoRuntime.queryInterface(XShapes.class, 
 						    xDrawPageSupplier.getDrawPage());
 	    // Add both shapes
 	    xShapes.add (xEllipse);
 	    xShapes.add (xRect);
	}
	
	// 	// Add paragraph break
	// 	doc.addParagraphBrake();
	
	// 	// Text
	// 	doc.getXText().insertString(doc.getXTextCursor(), 
	// 				    "This is a small Paragraph AFTER the graphics group", 
	// 				    false );
	

	// 	System.out.println("+++++++++++++++++++++++++++++++++++++++");
	// 	System.out.println("doc.getFrameWidth()="+doc.getFrameWidth());
	// 	System.out.println("+++++++++++++++++++++++++++++++++++++++");
	// 	printProperties("doc", doc.getXTextDocument());
	// 	System.out.println("+++++++++++++++++++++++++++++++++++++++");
	// 	printProperties("text", doc.getXText());
	// 	System.out.println("+++++++++++++++++++++++++++++++++++++++");
	// 	printProperties("desktop", doc.getXDesktop());
	// 	System.out.println("+++++++++++++++++++++++++++++++++++++++");

	// 	// Add paragraph break
	// 	doc.addParagraphBrake();

	doc.addParagraph("Drawing within a Text Document", "Heading 1");
	doc.addParagraph("Document Generated "+(new java.util.Date()).toString());
	doc.addParagraphBrake();
	doc.addParagraph(":-)");
	

	doc.addPageBreak();
	doc.addParagraph("Test Graph 1", "Heading 1");
	doc.addParagraph("Document Generated "+(new java.util.Date()).toString());
	doc.addParagraphBrake();

	//
	// Use DrawCache
	// 
	DrawCache cache = new DrawCache();

	// Create bg rect
	Object tbgrectobj = doc.getDocumentFactory().createInstance("com.sun.star.drawing.RectangleShape");
	XShape tbgrect = (XShape)UnoRuntime.queryInterface(XShape.class, tbgrectobj);
	tbgrect.setPosition( new Point(0, 0) );
	tbgrect.setSize(new Size(doc.getFrameWidth(), 13000));
	cache.addShape(tbgrectobj, tbgrect);
	cache.setShapeProperty(tbgrectobj, "FillColor", new Integer(0xFFFFFF));
	cache.setShapeProperty(tbgrectobj, "TextContourFrame", new Boolean(false));

	// Create rectangle 
	Object trect1obj = doc.getDocumentFactory().createInstance("com.sun.star.drawing.RectangleShape");
	XShape trect1 = (XShape)UnoRuntime.queryInterface(XShape.class, trect1obj);
	trect1.setPosition( new Point( 3000, 3000 ) );
	trect1.setSize(new Size(3000, 3000));
	cache.addShape(trect1obj, trect1);
	cache.setShapeProperty(trect1obj, "TextContourFrame", new Boolean(true));
	cache.setShapeProperty(trect1obj, "TextFitToSize", TextFitToSizeType.PROPORTIONAL );
	cache.setShapeProperty(trect1obj, "FillColor", new Integer(0xFFFFFF));
	cache.setShapeText(trect1obj, "Rect 1");

	// Create rectangle 2
 	Object trect2obj = doc.getDocumentFactory().createInstance("com.sun.star.drawing.RectangleShape");
	XShape trect2 = (XShape)UnoRuntime.queryInterface(XShape.class, trect2obj);
	trect2.setPosition( new Point( 14000, 4000 ) );
	trect2.setSize(new Size(1000, 1000));
	cache.addShape(trect2obj, trect2);
	cache.setShapeProperty(trect2obj, "FillColor", new Integer(0xFFFFFF));
	cache.setShapeProperty(trect2obj, "TextFitToSize", TextFitToSizeType.PROPORTIONAL );
	cache.setShapeProperty(trect2obj, "TextContourFrame", new Boolean(true));
	cache.setShapeText(trect2obj, "Rect 2");

	cache.addCache2Text(doc.getDocumentFactory(), doc.getXText());

	// Add paragraph breaks after graph
	doc.addParagraphBrake();
	doc.addParagraphBrake();

	// Text
	doc.addParagraph("This is a small Paragraph AFTER the CacheDraw");
	doc.addParagraph("Another use of...");

	doc.addParagraph("Use Case Diagram", "Heading 2");
	doc.addParagraph("This is a small Use Case diagram");
	doc.addParagraphBrake();

	//
	// Use DrawCache
	// 
	DrawCache cache2 = new DrawCache();

	// Create bg rect
	Object tbgrect2obj = doc.getDocumentFactory().createInstance("com.sun.star.drawing.RectangleShape");
	XShape tbgrect2 = (XShape)UnoRuntime.queryInterface(XShape.class, tbgrect2obj);
	tbgrect2.setPosition( new Point(0, 0) );
	tbgrect2.setSize(new Size(doc.getFrameWidth(), 13000));
	cache2.addShape(tbgrect2obj, tbgrect2);
	cache2.setShapeProperty(tbgrect2obj, "FillColor", new Integer(0xFFFFFF));
	cache2.setShapeProperty(tbgrect2obj, "TextContourFrame", new Boolean(false));

	XShape fromActor = cache2.addActorXShape(doc.getDocumentFactory(), 
						 1000, 1000,
						 500, 1500,
						 "Actor 1");

	XShape toActor = cache2.addActorXShape(doc.getDocumentFactory(), 
					       5000, 3000,
					       500, 1500,
					       "Actor 2");

	XShape ass1 = cache2.addAssociationXShape(doc.getDocumentFactory(), 
						   fromActor, 
						   toActor); 		    

	XShape uc1 = cache2.addUseCaseXShape(doc.getDocumentFactory(), 
					   10000, 4000,
					   2000, 1000,
					   "Use Case 1");

	XShape ass2 = cache2.addAssociationXShape(doc.getDocumentFactory(), 
						   uc1, 
						   toActor); 


	cache2.addCache2Text(doc.getDocumentFactory(), doc.getXText());

	// Add paragraph breaks after graph
	doc.addParagraphBrake();
	doc.addParagraphBrake();

	doc.addParagraph("Class Diagram", "Heading 2");
	doc.addParagraph("This is a small Class diagram");
	doc.addParagraphBrake();

	//
	// Use DrawCache
	// 
	DrawCache cache3 = new DrawCache();

	// Create bg rect
	Object tbgrect3obj = doc.getDocumentFactory().createInstance("com.sun.star.drawing.RectangleShape");
	XShape tbgrect3 = (XShape)UnoRuntime.queryInterface(XShape.class, tbgrect3obj);
	tbgrect3.setPosition( new Point(0, 0) );
	tbgrect3.setSize(new Size(doc.getFrameWidth(), 13000));
	cache3.addShape(tbgrect3obj, tbgrect3);
	cache3.setShapeProperty(tbgrect3obj, "FillColor", new Integer(0xFFFFFF));
	cache3.setShapeProperty(tbgrect3obj, "TextContourFrame", new Boolean(false));

	XShape superClass = cache3.addClassXShape(doc.getDocumentFactory(), 
					    4000, 1000, 
					    2000, 1000,
					    "Super class");

	XShape sub1 = cache3.addClassXShape(doc.getDocumentFactory(), 
				      3000, 6000, 
				      2000, 1000,
				      "Sub class 1");

	XShape sub15 = cache3.addClassXShape(doc.getDocumentFactory(), 
				      3000, 6500, 
				      2000, 1000,
				      "Sub class 1.5");

	XShape sub2 = cache3.addClassXShape(doc.getDocumentFactory(), 
				      8000, 6000, 
				      2000, 1000,
				      "Sub class 2");

 	XShape inheritance = 
 	    cache3.addInheritanceTriangleXShape(doc.getDocumentFactory(), 
 					  superClass);

 	XShape line1 = cache3.addRectLineXShape(doc.getDocumentFactory(), 
 					  sub1, 
 					  inheritance);

 	XShape line2 = cache3.addRectLineXShape(doc.getDocumentFactory(), 
 					  sub2, 
 					  inheritance);
	    

	cache3.addCache2Text(doc.getDocumentFactory(), doc.getXText());

	// Add paragraph breaks after graph
	doc.addParagraphBrake();
	doc.addParagraphBrake();

	doc.addParagraph("More Text...");

    }

    public static void printProperties(String name, java.lang.Object comp) 
	throws java.lang.Exception {
	XPropertySet xPageProps = Soffice.getXPropertySet(comp);
	XPropertySetInfo pageInfo = xPageProps.getPropertySetInfo();
	com.sun.star.beans.Property props[] = pageInfo.getProperties();
	for(int i=0;i<props.length;i++) {
	    com.sun.star.beans.Property prop = props[i];
	    System.out.println(name+"."+prop.Name+
			       ": handle="+prop.Handle+
			       ", Type="+prop.Type+
			       ", Attributes="+prop.Attributes);
	}
    }
    
}
