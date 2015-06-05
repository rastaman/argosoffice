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

import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.XPropertySet;
import com.sun.star.bridge.XUnoUrlResolver;
import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.container.XEnumeration;
import com.sun.star.container.XEnumerationAccess;
import com.sun.star.container.XIndexAccess;
import com.sun.star.drawing.DashStyle;
import com.sun.star.drawing.LineDash;
import com.sun.star.drawing.LineStyle;
import com.sun.star.drawing.XShape;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.frame.XDesktop;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.text.ControlCharacter;
import com.sun.star.text.XText;
import com.sun.star.text.XTextContent;
import com.sun.star.text.XTextCursor;
import com.sun.star.text.XTextDocument;
import com.sun.star.text.XTextRange;
import com.sun.star.text.XTextTable;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import com.sun.star.uno.XInterface;
import com.sun.star.uno.XNamingService;

/**
 * Facade to simplify Star/OpenOffice interaction
 * 
 *
 * @author <a href="mailto:matti@sun.com">Matti Pehrs</a>
 * @version $Id: Soffice.java,v 1.4 2005/10/31 06:51:25 rastaman Exp $
 */
public abstract class Soffice {

    public final static int COLOR_BLACK = 0x000000;
    public final static int COLOR_WHITE = 0xFFFFFF;

    /**
     * Connect (and get reference) to StarOffice NameService on specified Host/Port
     */
    public static XMultiServiceFactory connect( String host, int port )
	throws com.sun.star.uno.Exception,
	       com.sun.star.uno.RuntimeException, 
	       Exception {

        String sConnectionString = "uno:socket,host="+host+",port="+port+";urp;StarOffice.NamingService";
	return connect(sConnectionString);
    }

    /**
     * Connect (and get reference) to StarOffice Service
     */
    public static XMultiServiceFactory connect( String connectStr )
	throws com.sun.star.uno.Exception,
	       com.sun.star.uno.RuntimeException, 
	       Exception {
        // Get component context
	// com.sun.star.comp.helper.Bootstrap
        XComponentContext xcomponentcontext = Bootstrap.createInitialComponentContext(null);
        
        // initial serviceManager
        XMultiComponentFactory xLocalServiceManager =
	    xcomponentcontext.getServiceManager();
                
        // create a connector, so that it can contact the office
        Object  xUrlResolver  = 
	    xLocalServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", 
							   xcomponentcontext );
        XUnoUrlResolver urlResolver = 
	    (XUnoUrlResolver)UnoRuntime.queryInterface(XUnoUrlResolver.class, xUrlResolver );
        
        Object rInitialObject = urlResolver.resolve( connectStr );
        
        XNamingService rName = 
	    (XNamingService)UnoRuntime.queryInterface(XNamingService.class, rInitialObject );
        
        XMultiServiceFactory xMSF = null;
        if( rName != null ) {
            System.err.println( "got the remote naming service !" );
            Object rXsmgr = rName.getRegisteredObject("StarOffice.ServiceManager" );
            
            xMSF = 
		(XMultiServiceFactory)UnoRuntime.queryInterface( XMultiServiceFactory.class, rXsmgr );
        }
        
        return xMSF;
    }
    

    /**
     * Create new Empty StarOffice Text Document
     */
    public static XTextDocument openWriter(XMultiServiceFactory oMSF) 
	throws Exception {

	String doc = "private:factory/swriter";

	return openWriter(oMSF, doc);
    }

    /**
     * Open StarOffice Text Document
     */
    public static XTextDocument openWriter(XMultiServiceFactory oMSF, String doc) 
	throws Exception {
        
        
        //define variables
        XInterface oInterface;
        XDesktop oDesktop;
        XComponentLoader oCLoader;
        XTextDocument oDoc = null;
        XComponent aDoc = null;
        
        try {
            
            oInterface = (XInterface) oMSF.createInstance( "com.sun.star.frame.Desktop" );
            oDesktop = ( XDesktop ) UnoRuntime.queryInterface( XDesktop.class, oInterface );
            oCLoader = ( XComponentLoader ) UnoRuntime.queryInterface( XComponentLoader.class, oDesktop );
            PropertyValue [] szEmptyArgs = new PropertyValue [0];
	    
            aDoc = oCLoader.loadComponentFromURL(doc, "_blank", 0, szEmptyArgs );
            oDoc = (XTextDocument) UnoRuntime.queryInterface(XTextDocument.class, aDoc);

	    if(oDoc==null) {
		throw new java.io.IOException("Could not find document "+doc);
	    }
            
        } catch(Exception e) {
            e.printStackTrace();
            System.out.println(" Exception " + e);
            
        } // end of catch

        return oDoc;
    }//end of openWriter


    public static XComponent openDocXComponent(XMultiServiceFactory oMSF, String doc) 
	throws Exception {
        
        //define variables
        XInterface oInterface;
        XDesktop oDesktop;
        XComponentLoader oCLoader;
        XTextDocument oDoc = null;
        XComponent aDoc = null;
        
        try {
            
            oInterface = (XInterface) oMSF.createInstance( "com.sun.star.frame.Desktop" );
            oDesktop = ( XDesktop ) UnoRuntime.queryInterface( XDesktop.class, oInterface );
            oCLoader = ( XComponentLoader ) UnoRuntime.queryInterface( XComponentLoader.class, oDesktop );
            PropertyValue [] szEmptyArgs = new PropertyValue [0];
	    
            aDoc = oCLoader.loadComponentFromURL(doc, "_blank", 0, szEmptyArgs );

	    return aDoc;
            
        } catch(Exception e) {
            e.printStackTrace();
            System.out.println(" Exception " + e);
        } // end of catch

        return aDoc;
    }//end of openWriter


    //
    // Text Functions
    //

    public static void insertParagraphBreak(XText oText, XTextCursor oCursor) 
	throws Exception {
	oText.insertControlCharacter( oCursor, com.sun.star.text.ControlCharacter.PARAGRAPH_BREAK, false );
    }


    /**
     * Set the last entered paragraph style
     */
    public static void setLastParagraphStyle(XText oText, String styleName) 
	throws Exception {
	
	// Get All Paragraphs
	XEnumerationAccess xEnumerationAccess = null;
	xEnumerationAccess = 
	    (XEnumerationAccess)UnoRuntime.queryInterface(XEnumerationAccess.class, oText );

	// the enumeration contains all paragraph form the document
	XEnumeration xParagraphEnumeration = 
	    xEnumerationAccess.createEnumeration();
	
	// Get the Last Paragraph
	XTextContent xParagraph = null;
	while(xParagraphEnumeration.hasMoreElements()) {
	    // get the next paragraph
	    xParagraph = 
		(XTextContent) UnoRuntime.queryInterface(XTextContent.class, 
							 xParagraphEnumeration.nextElement());
	}
	if(xParagraph==null) return; // Ignore if paragraph could not be found
	
	// Get the Cursor
	XTextCursor xParaCursor = 
	    xParagraph.getAnchor().getText().createTextCursor();
	
	// Goto the start of the paragraph 
	// and select to the end of it
	xParaCursor.gotoStart( false );
	xParaCursor.gotoEnd( true );

	// Go through all parts of the paragraph
	XEnumerationAccess xParaEnumerationAccess = 
	    (XEnumerationAccess)UnoRuntime.queryInterface(XEnumerationAccess.class, xParagraph);
	XEnumeration xPortionEnumeration = 
	    xParaEnumerationAccess.createEnumeration();                

	while(xPortionEnumeration.hasMoreElements()) {

	    // Get the text range part of the paragraph
	    XTextRange xWord = 
		(XTextRange) UnoRuntime.queryInterface(XTextRange.class, 
						       xPortionEnumeration.nextElement());
	    // String sWordString = xWord.getString();
	    // System.out.println( "Content of the paragraph : " + sWordString );
	    
	    // Set the style
	    XPropertySet xPropertySet = 
		(XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, 
							 xWord );
	    // Set Paragraph Style 
	    xPropertySet.setPropertyValue("ParaStyleName", styleName);
	}	    
	
    }

    /**
     * Create a new Text Table
     */
    public static XTextTable createXTextTable(XTextDocument myDoc) 
	throws Exception {
	
        //getting MSF of the document
        XMultiServiceFactory oDocMSF = 
	    (XMultiServiceFactory) UnoRuntime.queryInterface(XMultiServiceFactory.class, myDoc );
	
	Object oInt = oDocMSF.createInstance("com.sun.star.text.TextTable");
	return (XTextTable) UnoRuntime.queryInterface(XTextTable.class,oInt);
    }

    /**
     * Get the PropertySet for a row in the Text Table
     */
    public static XPropertySet getXTextTableRowPropertySet(XTextTable TT, int rowNum) 
	throws Exception {

	// get Row Properties
	XIndexAccess theRows = TT.getRows();
	return (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, 
							theRows.getByIndex(rowNum));
    }

    /**
     * Get the PropertySet for the Text Table
     */
    public static XPropertySet getXTextTablePropertySet(XTextTable TT) 
	throws Exception {
	// get the property set of the text table
	return (XPropertySet)UnoRuntime.queryInterface( XPropertySet.class, TT );        
    }

    
    public static XPropertySet getXPropertySet(java.lang.Object obj) 
	throws Exception {
	return (XPropertySet)UnoRuntime.queryInterface(XPropertySet.class, 
						       obj);
    }
    
    /**
     * Set the Background Transparency
     */
    public static void setBackTransparent(XPropertySet props, boolean val) 
	throws Exception {
	props.setPropertyValue("BackTransparent", new Boolean(val));
    }

    /**
     * Set the Background Color
     */
    public static void setBackColor(XPropertySet props, int val) 
	throws Exception {
	props.setPropertyValue("BackColor", new Integer(val));
    }

    /**
     * Set the Character Color
     */
    public static void setCharColor(XPropertySet props, int val) 
	throws Exception {
	props.setPropertyValue("CharColor", new Integer(val));
    }

    /**
     * Set the Character Color
     */
    public static void setParaStyleName(XPropertySet props, String val) 
	throws Exception {
	props.setPropertyValue("ParaStyleName", val);
    }

    /**
     * Insert text value into text table cell
     */
    public static XPropertySet insertText2Cell(XTextTable TT, String CellName, String theText) {
	
        XText oTableText = (XText) UnoRuntime.queryInterface(XText.class, TT.getCellByName(CellName));

	if(oTableText == null) {
	    System.out.println("ERROR in Soffice.insertText2Cell(TT, '"+CellName+"', '"+theText+"')");
	    return null;
	}
        
        //create a cursor object
        XTextCursor oC = oTableText.createTextCursor();
        XPropertySet oTPS = (XPropertySet)UnoRuntime.queryInterface( XPropertySet.class, oC );
        
        // insert the Text
        oTableText.setString( "" + theText );

	return oTPS;
        
    } 
    

    /**
     * add text to a shape. the return value is the PropertySet
     * of the text range that has been added
     */
    public static XPropertySet addPortion(XShape xShape, 
					  String sText, 
					  boolean bNewParagraph )
	throws com.sun.star.lang.IllegalArgumentException
    {
	XText xText = (XText)
	    UnoRuntime.queryInterface( XText.class, xShape );
	
	XTextCursor xTextCursor = xText.createTextCursor();
	xTextCursor.gotoEnd(false);
	if(bNewParagraph == true) {
	    xText.insertControlCharacter( xTextCursor, ControlCharacter.PARAGRAPH_BREAK, false );
	    xTextCursor.gotoEnd( false );
	}
	XTextRange xTextRange = (XTextRange)
	    UnoRuntime.queryInterface( XTextRange.class, xTextCursor );
	xTextRange.setString( sText );
	xTextCursor.gotoEnd( true );
	XPropertySet xPropSet = (XPropertySet)
	    UnoRuntime.queryInterface( XPropertySet.class, xTextRange );
	return xPropSet;
    }


    public static void setDashedLine(XShape compAss) 
	throws java.lang.Exception {
	
	XPropertySet assProps = 
	    (XPropertySet)UnoRuntime.queryInterface(com.sun.star.beans.XPropertySet.class, 
						    compAss);
	assProps.setPropertyValue("LineStyle", LineStyle.DASH);
	LineDash dash = new LineDash();
	dash.Style = DashStyle.RECT;
	dash.Dashes = 1;
	dash.DashLen = 150;
	dash.Distance = 150;
	assProps.setPropertyValue("LineDash", dash);
    }
    
}
