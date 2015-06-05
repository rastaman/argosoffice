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
import org.apache.log4j.Logger;

import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.XPropertySet;
import com.sun.star.beans.XPropertySetInfo;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.frame.XDesktop;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.style.BreakType;
import com.sun.star.text.ControlCharacter;
import com.sun.star.text.XText;
import com.sun.star.text.XTextCursor;
import com.sun.star.text.XTextDocument;
import com.sun.star.text.XTextFrame;
import com.sun.star.text.XTextFramesSupplier;
import com.sun.star.uno.AnyConverter;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XInterface;

/**
 * Represents a Star/OpenOffice Draw Document
 * 
 * @author Matti Pehrs
 * @version $Id: SWriterDoc.java,v 1.4 2008/02/05 10:33:08 rastaman Exp $
 */
public class SWriterDoc {

	private Logger log = Logger.getLogger(SWriterDoc.class);
	
    public SWriterDoc() {
    }

    // define variables
    XInterface desktopObject;

    XDesktop oDesktop;

    XComponentLoader oCLoader;

    XComponent aDoc = null;

    XTextDocument oDoc = null;

    XText oText = null;

    XTextCursor oCursor = null;

    XMultiServiceFactory xDocumentFactory = null;

    XMultiServiceFactory oMSF;

    public XMultiServiceFactory getXMultiServiceFactory() {
        return oMSF;
    }

    public XComponentLoader getXComponentLoader() {
        return oCLoader;
    }

    public XInterface getXDesktopObject() {
        return desktopObject;
    }

    public XDesktop getXDesktop() {
        return oDesktop;
    }

    public XMultiServiceFactory getDocumentFactory() {
        return xDocumentFactory;
    }

    public XTextCursor getXTextCursor() {
        return oCursor;
    }

    public XText getXText() {
        return oText;
    }

    public XComponent getXComponent() {
        return aDoc;
    }

    public XTextDocument getXTextDocument() {
        return oDoc;
    }

    public void connect() throws java.lang.Exception {
        connect("localhost", 8100, "private:factory/swriter");
    }

    public void connect(String host, int port) throws java.lang.Exception {
        connect(host, port, "private:factory/swriter");
    }

    public void connect(String templateDoc) throws java.lang.Exception {
        connect("localhost", 8100, templateDoc);
    }

    public void connect(String host, int port, String templateDoc)
            throws java.lang.Exception {

        try {
            oMSF = Soffice.connect(host, port);

            desktopObject = (XInterface) oMSF
                    .createInstance("com.sun.star.frame.Desktop");
            oDesktop = (XDesktop) UnoRuntime.queryInterface(XDesktop.class,
                    desktopObject);
            oCLoader = (XComponentLoader) UnoRuntime.queryInterface(
                    XComponentLoader.class, oDesktop);
            PropertyValue[] szEmptyArgs = new PropertyValue[0];

            log.error("Loading template "+templateDoc);
            aDoc = oCLoader.loadComponentFromURL(templateDoc, "_blank", 0,
                    szEmptyArgs);
            oDoc = (XTextDocument) UnoRuntime.queryInterface(
                    XTextDocument.class, aDoc);

            if (oDoc == null) {
                throw new java.io.IOException("Could not find document "
                        + templateDoc);
            }

            // getting the text object
            oText = oDoc.getText();
            // create a cursor object
            oCursor = oText.createTextCursor();

            xDocumentFactory = (XMultiServiceFactory) UnoRuntime
                    .queryInterface(XMultiServiceFactory.class, oDoc);

        } catch (java.lang.Exception e) {
            e.printStackTrace();
            System.out.println(" Exception " + e);

        } // end of catch

    }

    public static void printProperties(String name, XPropertySet xPageProps)
            throws java.lang.Exception {
        XPropertySetInfo pageInfo = xPageProps.getPropertySetInfo();
        com.sun.star.beans.Property props[] = pageInfo.getProperties();
        for (int i = 0; i < props.length; i++) {
            com.sun.star.beans.Property prop = props[i];
            System.out.println(name + "." + prop.Name + ": handle="
                    + prop.Handle + ", Type=" + prop.Type + ", Attributes="
                    + prop.Attributes);
        }
    }

    /**
     * Extract the first frames width
     */
    public int getFrameWidth() {
        // Get Frame info
        if (1 == 0) {
            try {
                XTextFramesSupplier frameSupplier = (XTextFramesSupplier) UnoRuntime
                        .queryInterface(XTextFramesSupplier.class,
                                getXComponent());
                System.out.println("frameSupplier=" + frameSupplier);
                com.sun.star.container.XNameAccess framesAccess = frameSupplier
                        .getTextFrames();
                String names[] = framesAccess.getElementNames();
                for (int i = 0; i < names.length; i++) {
                    String name = names[i];
                    System.out.println("frame[" + i + "].name=" + name);
                    com.sun.star.uno.Any any = (com.sun.star.uno.Any) framesAccess
                            .getByName(name);
                    XTextFrame frame = (XTextFrame) any.getObject();
                    XPropertySet frameProps = Soffice.getXPropertySet(frame);
                    System.out.println("  width="
                            + AnyConverter.toInt(frameProps
                                    .getPropertyValue("FrameWidthAbsolute")));
                    return AnyConverter.toInt(frameProps
                            .getPropertyValue("FrameWidthAbsolute"));
                }

            } catch (java.lang.Exception ex) {
                ex.printStackTrace();
            }
            return -1;
        }

        return 23310; // FIXME...
    }

    public void addParagraph(String text) throws java.lang.Exception {
        oCursor.gotoEnd(false);
        addParagraph(text, "Default");
    }

    public void addParagraph(String text, String style)
            throws java.lang.Exception {
        oCursor.gotoEnd(false);
        oText.insertString(oCursor, text, false);
        if (style!=null)
            Soffice.setLastParagraphStyle(oText, style);
        Soffice.insertParagraphBreak(oText, oCursor);      
    }

    public void addParagraphBrake() throws java.lang.Exception {
        // Insert a paragraph break into the document (not the frame)
        oText.insertControlCharacter(oCursor, ControlCharacter.PARAGRAPH_BREAK,
                false);
    }

    public void addPageBreak() throws java.lang.Exception {
        // Insert a paragraph break into the document (not the frame)
        // oText.insertControlCharacter (oCursor, ControlCharacter.PAGE_BREAK,
        // false );

        XPropertySet xCursorProps = (XPropertySet) UnoRuntime.queryInterface(
                XPropertySet.class, oCursor);
        // call setPropertyValue, passing in a Float object
        xCursorProps.setPropertyValue("BreakType", BreakType.PAGE_BEFORE);
    }

}
