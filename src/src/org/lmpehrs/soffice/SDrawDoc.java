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
import java.util.Iterator;
import java.util.List;
import java.util.Vector;

import org.apache.log4j.Logger;

import com.sun.star.awt.Point;
import com.sun.star.awt.Size;
import com.sun.star.beans.XPropertySet;
import com.sun.star.container.XNamed;
import com.sun.star.drawing.ConnectorType;
import com.sun.star.drawing.DashStyle;
import com.sun.star.drawing.FillStyle;
import com.sun.star.drawing.LineDash;
import com.sun.star.drawing.LineStyle;
import com.sun.star.drawing.PolyPolygonBezierCoords;
import com.sun.star.drawing.PolygonFlags;
import com.sun.star.drawing.TextFitToSizeType;
import com.sun.star.drawing.TextHorizontalAdjust;
import com.sun.star.drawing.TextVerticalAdjust;
import com.sun.star.drawing.XDrawPage;
import com.sun.star.drawing.XDrawPages;
import com.sun.star.drawing.XDrawPagesSupplier;
import com.sun.star.drawing.XShape;
import com.sun.star.drawing.XShapeGrouper;
import com.sun.star.drawing.XShapes;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.style.ParagraphAdjust;
import com.sun.star.uno.AnyConverter;
import com.sun.star.uno.UnoRuntime;

/**
 * Represents a Star/OpenOffice Draw Document
 *
 * @author  Matti Pehrs
 * @version $Id: SDrawDoc.java,v 1.4 2008/02/05 10:33:08 rastaman Exp $
 */
public class SDrawDoc {

    private Logger LOG = Logger.getLogger(SDrawDoc.class);
    
    XMultiServiceFactory msf = null;
    XMultiServiceFactory xDocumentFactory = null;
    XComponent xDrawComponent = null;
    XDrawPagesSupplier xDrawPagesSupplier = null;
    XDrawPages xDrawPages = null;
    Object drawPage = null;
    XDrawPage xDrawPage = null;
    XPropertySet xPageProps = null;
    
    int pageWidth;
    int pageHeight;
    int pageBorderTop;
    int pageBorderBottom;
    int pageBorderLeft;
    int pageBorderRight;
    int drawWidth;
    int drawHeight;
    int horCenter;

    int virtualWidth;
    int virtualHeight;

    float dx = 1.0f;
    float dy = 1.0f;

    public SDrawDoc() {
    }

    public int getDrawWidth() { return drawWidth;}
    public int getDrawHeight() { return drawHeight;}

    public void setDX(float val) { dx = val;}
    public void setDY(float val) { dy = val;}

    /**
     * Should be called AFTER connect()
     */
    public void setVirtualWidth(int virtualWidth) {
	this.virtualWidth = virtualWidth;
	dx = (float)drawWidth / (float)virtualWidth;
	// System.out.println("virtualWidth="+virtualWidth+", drawWidth="+drawWidth+", dx="+dx);
    }

    /**
     * Should be called AFTER connect()
     */
    public void setVirtualHeight(int virtualHeight) {
	this.virtualHeight = virtualHeight;
	dy = (float)drawHeight / (float)virtualHeight;
	//System.out.println("virtualHeight="+virtualHeight+", drawHeight="+drawHeight+", dy="+dy);
    }
    
    /**
     * Connect to the StarOffice server with a template as parameter.
     * @param templateDoc
     * @throws java.lang.Exception
     */
    public void connect(String templateDoc) throws java.lang.Exception {
	connect("localhost", 8100, templateDoc);
    }

    /**
     * Connect to a StarOffice host to the given port, with the given doc.
     * @param host
     * @param port
     * @param templateDoc
     * @throws java.lang.Exception
     */
    public void connect(String host, int port, String templateDoc) throws java.lang.Exception {

	msf = Soffice.connect(host, port);
	
    LOG.error("opening "+templateDoc);
	xDrawComponent = Soffice.openDocXComponent(msf, templateDoc);
	// xDrawComponent = Soffice.openDocXComponent(msf, "private:factory/sdraw");
	
	xDocumentFactory = 
	    (XMultiServiceFactory)UnoRuntime.queryInterface(XMultiServiceFactory.class, 
							    xDrawComponent);	
        // get draw page by index
	xDrawPagesSupplier = 
	    (XDrawPagesSupplier)UnoRuntime.queryInterface(XDrawPagesSupplier.class, 
							  xDrawComponent );
	xDrawPages = xDrawPagesSupplier.getDrawPages();
	drawPage = xDrawPages.getByIndex(0);
	xDrawPage = 
	    (XDrawPage)UnoRuntime.queryInterface(XDrawPage.class, 
						 drawPage);
	xPageProps = Soffice.getXPropertySet(xDrawPage);
	
	pageWidth = AnyConverter.toInt(xPageProps.getPropertyValue("Width"));
	pageHeight = AnyConverter.toInt(xPageProps.getPropertyValue("Height"));
	pageBorderTop = AnyConverter.toInt(xPageProps.getPropertyValue("BorderTop"));
	pageBorderBottom = AnyConverter.toInt(xPageProps.getPropertyValue("BorderBottom"));
	pageBorderLeft = AnyConverter.toInt(xPageProps.getPropertyValue("BorderLeft"));
	pageBorderRight = AnyConverter.toInt(xPageProps.getPropertyValue("BorderRight"));
	drawWidth = pageWidth - pageBorderLeft - pageBorderRight;
	drawHeight = pageHeight - pageBorderTop - pageBorderBottom;
	horCenter = pageBorderLeft + drawWidth / 2;
    }

    /**
     * Draw a UseCase in StarOffice
     * @param x
     * @param y
     * @param width
     * @param height
     * @param name
     * @return Xshape the XShape object representing this UseCase in StarOffice
     * @throws java.lang.Exception
     */
    public XShape drawUseCase(int x, int y, 
			      int width, int height,
			      String name) 
	throws java.lang.Exception {

	//System.out.println("UseCase("+x+", "+y+", "+width+", "+height+", "+name+")");
	x = (int)(dx * x);	
	y = (int)(dy * y);

	width = (int)(dx * width);
	height = (int)(dx * height);

	//System.out.println("  -->["+x+", "+y+", "+width+", "+height+"]");

	// 
	// UseCase Object 
	//
	Object shape = xDocumentFactory.createInstance("com.sun.star.drawing.EllipseShape");
	XShape xShape = 
	    (XShape)UnoRuntime.queryInterface(XShape.class, shape);
	xShape.setPosition(new Point(x, y));
	xShape.setSize(new Size(width, height));

	// ucXShapes.add(xShape);
	xDrawPage.add(xShape);

	XPropertySet xShapePropSet = Soffice.getXPropertySet(xShape);
	
	// xShapePropSet.setPropertyValue("TextFitToSize", TextFitToSizeType.PROPORTIONAL );
	xShapePropSet.setPropertyValue("FillStyle", FillStyle.NONE);
	// 	xShapePropSet.setPropertyValue("TextLeftDistance",  new Integer(500) );
	// 	xShapePropSet.setPropertyValue("TextRightDistance", new Integer(500) );
	// 	xShapePropSet.setPropertyValue("TextUpperDistance", new Integer(500) );
	// 	xShapePropSet.setPropertyValue("TextLowerDistance", new Integer(500) );

	XPropertySet xTextPropSet = Soffice.addPortion(xShape, name, false );
	// xTextPropSet.setPropertyValue( "ParaAdjust", ParagraphAdjust.CENTER );
	//xTextPropSet.setPropertyValue("CharColor",  new Integer( SofficeDialog.COLOR_SUN2 ) );
	xTextPropSet.setPropertyValue("CharColor",  new Integer(0x000000));
	xTextPropSet.setPropertyValue("CharHeight",  new Integer(10));
 	xTextPropSet.setPropertyValue("CharFontName", "Arial");

	// 	xTextPropSet = Soffice.addPortion(xShape, name, true );
	// 	xTextPropSet.setPropertyValue( "ParaAdjust", ParagraphAdjust.CENTER );
	// 	xTextPropSet.setPropertyValue("CharColor",  new Integer( SofficeDialog.COLOR_SUN2 ) );
	//  	xTextPropSet.setPropertyValue("CharFontName", "Arial");
	
	return xShape;
    }

    public XShape drawText(int x, int y, 
			   int width, int height,
			   String text,
			   String fontFamily,
			   int fontSize) 
	throws java.lang.Exception {
	return drawText( x, y, 
			 width, height,
			 text,
			 fontSize);
    }
    
    public XShape drawText(int x, int y, 
			   int width, int height,
			   String text,
			   int fontSize)
	throws java.lang.Exception {

	if(text == null) return null;
	if(text.length() == 0) return null;

	x = (int)(dx * x);	
	y = (int)(dy * y);

	width = (int)(dx * width);
	height = (int)(dy * height);
	
	//
	// Class Rectangle
	// 
	Object shape2 = xDocumentFactory.createInstance("com.sun.star.drawing.RectangleShape");
	XShape xShape2 = 
	    (XShape)UnoRuntime.queryInterface(XShape.class, shape2);
	xShape2.setPosition(new Point(x, y));
	xShape2.setSize(new Size(width, height));  

	xDrawPage.add(xShape2);

	XPropertySet xShapePropSet2 = Soffice.getXPropertySet(xShape2);
	
	if(drawTextPropotionalText) {
	     xShapePropSet2.setPropertyValue("TextFitToSize", TextFitToSizeType.PROPORTIONAL ); // FIXME: Properties
	}
	xShapePropSet2.setPropertyValue("FillStyle", FillStyle.NONE);
	xShapePropSet2.setPropertyValue("FillColor", new Integer(0xFFFFFF));
	xShapePropSet2.setPropertyValue("LineStyle", LineStyle.NONE);
	xShapePropSet2.setPropertyValue("TextLeftDistance",  new Integer(100) ); // FIXME: Properties
	xShapePropSet2.setPropertyValue("TextRightDistance", new Integer(100) ); // FIXME: Properties
	xShapePropSet2.setPropertyValue("TextUpperDistance", new Integer(20) );	// FIXME: Properties
	xShapePropSet2.setPropertyValue("TextLowerDistance", new Integer(20) );	// FIXME: Properties

	XPropertySet xTextPropSet2 = Soffice.addPortion(xShape2, text, false );
	// 	xTextPropSet2.setPropertyValue("ParaAdjust", ParagraphAdjust.CENTER );
	xTextPropSet2.setPropertyValue("CharColor", new Integer(0)); 
	xTextPropSet2.setPropertyValue("CharHeight", new Integer(drawTextfontSize));
	xTextPropSet2.setPropertyValue("CharFontName", drawTextFontFamily); 

	return xShape2;
    }

    boolean drawTextPropotionalText = false;
    int drawTextfontSize = 8;
    String drawTextFontFamily = "Dialog";

    public void setDrawTextFontFamily(String val) {
	drawTextFontFamily = val;
    }

    public void setDrawTextPropotionalText(boolean val) {
	drawTextPropotionalText = val;
    }

    public void setDrawTextfontSize(int val) {
	drawTextfontSize = val;
    }

    public double dist(int x0, int y0, int x1, int y1) {
	double dx, dy;
	dx = (double)(x0-x1);
	dy = (double)(y0-y1);
	return Math.sqrt(dx*dx+dy*dy);
    }
    
    public double dist(double dx, double dy) {
	return Math.sqrt(dx*dx+dy*dy);
    }

    public void drawDiamondArrowHead(int xFrom, int yFrom, 
				     int xTo, int yTo,
				     int fillColor) 
	throws java.lang.Exception {
	double denom, x, y, ddx, ddy, cos, sin;
	int arrow_width = (int)(dx * 7);
	int arrow_height = (int)(dx * 12);

	xTo = (int)(dx * xTo);	
	xFrom = (int)(dx * xFrom);	
	yTo = (int)(dx * yTo);	
	yFrom = (int)(dx * yFrom);	
	
	ddx   	= (double)(xTo - xFrom);
	ddy   	= (double)(yTo - yFrom);
	denom 	= dist(ddx, ddy);
	if (denom == 0) return;
	
	cos = (arrow_height/2)/denom;
	sin = arrow_width /denom;
	x   = xTo - cos*ddx;
	y   = yTo - cos*ddy;
	int x1  = (int)(x - sin*ddy);
	int y1  = (int)(y + sin*ddx);
	int x2  = (int)(x + sin*ddy);
	int y2  = (int)(y - sin*ddx);

	java.awt.Point end = new java.awt.Point(xTo, yTo);
	java.awt.Point start = new java.awt.Point(xFrom, xFrom);
	java.awt.Point topPoint = pointAlongLine(end, start, arrow_height);
	
	Object xObj = xDocumentFactory.createInstance( "com.sun.star.drawing.ClosedBezierShape" );
	XShape polyShape = (XShape)UnoRuntime.queryInterface(XShape.class, xObj );

	xDrawPage.add(polyShape);

	XPropertySet xShapeProperties = 
	    (XPropertySet)UnoRuntime.queryInterface(XPropertySet.class, 
						    polyShape);

	PolyPolygonBezierCoords aCoords = new PolyPolygonBezierCoords();
	// allocating the outer sequence
	aCoords.Coordinates = new Point[1][ ];	
	aCoords.Flags = new PolygonFlags[1][ ];

	int nPointCount = 4;
	Point[]		pPolyPoints     = new Point[ nPointCount ];
	PolygonFlags[]	pPolyFlags	= new PolygonFlags[ nPointCount ];
	
	for(int n = 0; n < nPointCount; n++ ) {
	    pPolyPoints[ n ] = new Point();
	}
	
	// 	Polygon diamond;
	// 	diamond = new Polygon();
	// 	diamond.addPoint(xTo, yTo);
	// 	diamond.addPoint(x1, y1);
	// 	diamond.addPoint(topPoint.x, topPoint.y);
	// 	diamond.addPoint(x2, y2);
	
	pPolyPoints[ 0 ].X = xTo;
	pPolyPoints[ 0 ].Y = yTo;
	pPolyFlags[ 0 ] = PolygonFlags.NORMAL;
	pPolyPoints[ 1 ].X = x1;
	pPolyPoints[ 1 ].Y = y1;
	pPolyFlags[ 1 ] = PolygonFlags.NORMAL;
	pPolyPoints[ 2 ].X = topPoint.x;
	pPolyPoints[ 2 ].Y = topPoint.y;
	pPolyFlags[ 2 ] = PolygonFlags.NORMAL;
	pPolyPoints[ 3 ].X = x2;
	pPolyPoints[ 3 ].Y = y2;
	pPolyFlags[ 3 ] = PolygonFlags.NORMAL;
	//
	aCoords.Coordinates[0] = pPolyPoints;
	aCoords.Flags[0]       = pPolyFlags;
	//
	xShapeProperties.setPropertyValue( "PolyPolygonBezier", aCoords );
	xShapeProperties.setPropertyValue("LineColor", new Integer(0x000000));
	xShapeProperties.setPropertyValue("FillColor", new Integer(fillColor));
	// move the shape to the front by changing the ZOrder
	xShapeProperties.setPropertyValue( "ZOrder", new Integer(1000000) );

	
    }

    public void drawTriangleArrowHead(int xFrom, int yFrom, 
				      int xTo, int yTo) 
	throws java.lang.Exception {
	double denom, x, y, ddx, ddy, cos, sin;
	int arrow_width = (int)(dx * 7);
	int arrow_height = (int)(dx * 12);

	xTo = (int)(dx * xTo);	
	xFrom = (int)(dx * xFrom);	
	yTo = (int)(dx * yTo);	
	yFrom = (int)(dx * yFrom);	
	
	ddx   	= (double)(xTo - xFrom);
	ddy   	= (double)(yTo - yFrom);
	denom 	= dist(ddx, ddy);
	if (denom <= 0.01) return;
	
	cos = arrow_height / denom;
	sin = arrow_width / denom;
	x   = xTo - cos*ddx;
	y   = yTo - cos*ddy;
	int x1  = (int)(x - sin*ddy);
	int y1  = (int)(y + sin*ddx);
	int x2  = (int)(x + sin*ddy);
	int y2  = (int)(y - sin*ddx);

	//     triangle = new Polygon();
	//     triangle.addPoint(xTo, yTo);
	//     triangle.addPoint(x1, y1);
	//     triangle.addPoint(x2, y2);
	
	// DX/DY
	// 	x1 = (int)(dx * x1);	
	// 	x2 = (int)(dx * x2);	
	// 	y1 = (int)(dy * y1);
	// 	y2 = (int)(dy * y2);
	
	Object xObj = xDocumentFactory.createInstance( "com.sun.star.drawing.ClosedBezierShape" );
	XShape polyShape = (XShape)UnoRuntime.queryInterface(XShape.class, xObj );

	xDrawPage.add(polyShape);

	XPropertySet xShapeProperties = 
	    (XPropertySet)UnoRuntime.queryInterface(XPropertySet.class, 
						    polyShape);

	PolyPolygonBezierCoords aCoords = new PolyPolygonBezierCoords();
	// allocating the outer sequence
	aCoords.Coordinates = new Point[1][ ];	
	aCoords.Flags = new PolygonFlags[1][ ];

	int nPointCount = 3;
	Point[]		pPolyPoints     = new Point[ nPointCount ];
	PolygonFlags[]	pPolyFlags	= new PolygonFlags[ nPointCount ];
	
	for(int n = 0; n < nPointCount; n++ ) {
	    pPolyPoints[ n ] = new Point();
	}
	
	//     triangle.addPoint(xTo, yTo);
	//     triangle.addPoint(x1, y1);
	//     triangle.addPoint(x2, y2);
	
	pPolyPoints[ 0 ].X = xTo;
	pPolyPoints[ 0 ].Y = yTo;
	pPolyFlags[ 0 ] = PolygonFlags.NORMAL;
	pPolyPoints[ 1 ].X = x1;
	pPolyPoints[ 1 ].Y = y1;
	pPolyFlags[ 1 ] = PolygonFlags.NORMAL;
	pPolyPoints[ 2 ].X = x2;
	pPolyPoints[ 2 ].Y = y2;
	pPolyFlags[ 2 ] = PolygonFlags.NORMAL;
	//
	aCoords.Coordinates[0] = pPolyPoints;
	aCoords.Flags[0]       = pPolyFlags;
	//
	xShapeProperties.setPropertyValue( "PolyPolygonBezier", aCoords );
	// White fill
	xShapeProperties.setPropertyValue("FillColor", new Integer(0xFFFFFF));
	// move the shape to the front by changing the ZOrder
	xShapeProperties.setPropertyValue( "ZOrder", new Integer(1000000) );

  }

    public void drawHalfTriangleArrowHead(int xFrom, int yFrom, 
					  int xTo, int yTo) 
	throws java.lang.Exception {
	double denom, x, y, ddx, ddy, cos, sin;
	int arrow_width = (int)(dx * 7);
	int arrow_height = (int)(dx * 12);

	xTo = (int)(dx * xTo);	
	xFrom = (int)(dx * xFrom);	
	yTo = (int)(dx * yTo);	
	yFrom = (int)(dx * yFrom);	
	
	ddx   	= (double)(xTo - xFrom);
	ddy   	= (double)(yTo - yFrom);
	denom 	= dist(ddx, ddy);
	if (denom == 0) return;

	cos = arrow_height / denom;
	sin = arrow_width / denom;
	x   = xTo - cos*dx;
	y   = yTo - cos*dy;
	int x1  = (int)(x - sin*ddy);
	int y1  = (int)(y + sin*ddx);
	int x2  = (int)(x + sin*ddy);
	int y2  = (int)(y - sin*ddx);

	// DX/DY
// 	x1 = (int)(dx * x1);	
// 	x2 = (int)(dx * x2);	
// 	y1 = (int)(dy * y1);
// 	y2 = (int)(dy * y2);
	
	// 	triangle = new Polygon();
	// 	triangle.addPoint(xTo, yTo);
	// 	triangle.addPoint(xFrom, yFrom);
	// 	triangle.addPoint(x2, y2);
	// 	g.setColor(arrowFillColor);
	// 	g.fillPolygon(triangle);
	// 	g.setColor(arrowLineColor);
	// 	g.drawPolygon(triangle);

	Object xObj = xDocumentFactory.createInstance( "com.sun.star.drawing.ClosedBezierShape" );
	XShape polyShape = (XShape)UnoRuntime.queryInterface(XShape.class, xObj );

	xDrawPage.add(polyShape);

	XPropertySet xShapeProperties = 
	    (XPropertySet)UnoRuntime.queryInterface(XPropertySet.class, 
						    polyShape);

	PolyPolygonBezierCoords aCoords = new PolyPolygonBezierCoords();
	// allocating the outer sequence
	aCoords.Coordinates = new Point[1][ ];	
	aCoords.Flags = new PolygonFlags[1][ ];

	int nPointCount = 3;
	Point[]		pPolyPoints     = new Point[ nPointCount ];
	PolygonFlags[]	pPolyFlags	= new PolygonFlags[ nPointCount ];
	
	for(int n = 0; n < nPointCount; n++ ) {
	    pPolyPoints[ n ] = new Point();
	}
	
	// 	triangle.addPoint(xTo, yTo);
	// 	triangle.addPoint(xFrom, yFrom);
	// 	triangle.addPoint(x2, y2);
	
	pPolyPoints[ 0 ].X = xTo;
	pPolyPoints[ 0 ].Y = yTo;
	pPolyFlags[ 0 ] = PolygonFlags.NORMAL;
	pPolyPoints[ 1 ].X = xFrom;
	pPolyPoints[ 1 ].Y = yFrom;
	pPolyFlags[ 1 ] = PolygonFlags.NORMAL;
	pPolyPoints[ 2 ].X = x2;
	pPolyPoints[ 2 ].Y = y2;
	pPolyFlags[ 2 ] = PolygonFlags.NORMAL;
	//
	aCoords.Coordinates[0] = pPolyPoints;
	aCoords.Flags[0]       = pPolyFlags;
	//
	xShapeProperties.setPropertyValue( "PolyPolygonBezier", aCoords );
	// White fill
	xShapeProperties.setPropertyValue("FillColor", new Integer(0xFFFFFF));
	// move the shape to the front by changing the ZOrder
	xShapeProperties.setPropertyValue( "ZOrder", new Integer(1000000) );


  }

    public int getLineLength(java.awt.Point one, java.awt.Point two) {
	int dxdx = (two.x - one.x) * (two.x - one.x);
	int dydy = (two.y - one.y) * (two.y - one.y);
	return (int) Math.sqrt(dxdx + dydy);
    }
    
    public java.awt.Point pointAlongLine(java.awt.Point one, java.awt.Point two, int dist) {
	int len = getLineLength(one, two);
	int p = dist;
	if (len == 0) return one;
	return new java.awt.Point(one.x + ((two.x - one.x) * p) / len,
				  one.y + ((two.y - one.y) * p) / len);
    }

    public void drawGreaterArrowHead(int xFrom, int yFrom, 
				     int xTo, int yTo) 
	throws java.lang.Exception {
	double denom, x, y, dx, dy, cos, sin;
	int arrow_width = 7, arrow_height = 12;
	
	dx   	= (double)(xTo - xFrom);
	dy   	= (double)(yTo - yFrom);
	denom 	= dist(dx, dy);
	if (denom == 0) return;
	
	cos = arrow_height/denom;
	sin = arrow_width /denom;
	x   = xTo - cos*dx;
	y   = yTo - cos*dy;
	int x1  = (int)(x - sin*dy);
	int y1  = (int)(y + sin*dx);
	int x2  = (int)(x + sin*dy);
	int y2  = (int)(y - sin*dx);
	
	drawLine(x1, y1, xTo, yTo);
	drawLine(x2, y2, xTo, yTo);
  }


    public XShape drawPoly(List ppoints) // [java.awt.Point]
	throws java.lang.Exception {

	Vector points = new Vector();

	for(Iterator it = ppoints.iterator();
	    it.hasNext();) {
	    // 	x1 = (int)(dx * x1);	
	    // 	y1 = (int)(dy * y1);
	    java.awt.Point p = (java.awt.Point)it.next();
	    points.add(new java.awt.Point((int)(dx * p.getX()),
					  (int)(dy * p.getY())));
	}

	// 	triangle = new Polygon();
	// 	triangle.addPoint(xTo, yTo);
	// 	triangle.addPoint(xFrom, yFrom);
	// 	triangle.addPoint(x2, y2);
	// 	g.setColor(arrowFillColor);
	// 	g.fillPolygon(triangle);
	// 	g.setColor(arrowLineColor);
	// 	g.drawPolygon(triangle);

	Object xObj = xDocumentFactory.createInstance( "com.sun.star.drawing.ClosedBezierShape" );
	XShape polyShape = (XShape)UnoRuntime.queryInterface(XShape.class, xObj );

	xDrawPage.add(polyShape);

	XPropertySet xShapeProperties = 
	    (XPropertySet)UnoRuntime.queryInterface(XPropertySet.class, 
						    polyShape);

	PolyPolygonBezierCoords aCoords = new PolyPolygonBezierCoords();
	// allocating the outer sequence
	aCoords.Coordinates = new Point[1][ ];	
	aCoords.Flags = new PolygonFlags[1][ ];

	int nPointCount = points.size();
	Point[]		pPolyPoints     = new Point[ nPointCount ];
	PolygonFlags[]	pPolyFlags	= new PolygonFlags[ nPointCount ];
	
	int n = 0;
	for(Iterator it = points.iterator();
	    it.hasNext();n++) {
	    java.awt.Point p = (java.awt.Point)it.next();
	    pPolyPoints[ n ] = new Point();
	    pPolyPoints[ n ].X = (int)p.getX();
	    pPolyPoints[ n ].Y = (int)p.getY();
	    pPolyFlags[ n ] = PolygonFlags.NORMAL;
	}
	//
	aCoords.Coordinates[0] = pPolyPoints;
	aCoords.Flags[0]       = pPolyFlags;
	//
	xShapeProperties.setPropertyValue( "PolyPolygonBezier", aCoords );
	// White fill
	xShapeProperties.setPropertyValue("FillColor", new Integer(0xFFFFFF));
	// move the shape to the front by changing the ZOrder
	xShapeProperties.setPropertyValue( "ZOrder", new Integer(1000000) );

	return polyShape;
    }
    

    public XShape drawRect(int x, int y, 
			   int width, int height) 
	throws java.lang.Exception {

	x = (int)(dx * x);	
	y = (int)(dy * y);

	width = (int)(dx * width);
	height = (int)(dy * height);
	
	//
	// Class Rectangle
	// 
	Object shape2 = xDocumentFactory.createInstance("com.sun.star.drawing.RectangleShape");
	XShape xShape2 = 
	    (XShape)UnoRuntime.queryInterface(XShape.class, shape2);
	xShape2.setPosition(new Point(x, y));
	xShape2.setSize(new Size(width, height));  

	xDrawPage.add(xShape2);

	XPropertySet xShapePropSet2 = Soffice.getXPropertySet(xShape2);
	
	xShapePropSet2.setPropertyValue("FillColor", new Integer(0xFFFFFF));

	return xShape2;
    }

    public XShape drawLine(int x1, int y1, int x2, int y2) 
	throws java.lang.Exception {
	return drawLine( x1,  y1,  x2,  y2, false);
    }

    public XShape drawLine(int x1, int y1, int x2, int y2, boolean dashedLine) 
	throws java.lang.Exception {

	// System.out.println("line["+x1+","+y1+"]-["+x2+", "+y2+"]");

	x1 = (int)(dx * x1);	
	x2 = (int)(dx * x2);	
	y1 = (int)(dy * y1);
	y2 = (int)(dy * y2);

	// System.out.println("  => line["+x1+","+y1+"]-["+x2+", "+y2+"]");

	Object lineShape = xDocumentFactory.createInstance("com.sun.star.drawing.LineShape");
	XShape lineXShape = 
	    (XShape)UnoRuntime.queryInterface(XShape.class, lineShape);
	lineXShape.setPosition(new Point(x1, y1));
	lineXShape.setSize(new Size(x2 - x1, y2 - y1));  
	xDrawPage.add(lineXShape);
	XPropertySet lineXShapePropSet = Soffice.getXPropertySet(lineXShape);
	lineXShapePropSet.setPropertyValue("LineColor", new Integer(0));
	if(dashedLine) {
	    lineXShapePropSet.setPropertyValue("LineStyle", LineStyle.DASH);
	    LineDash dash = new LineDash();
	    dash.Style = DashStyle.RECT;
	    dash.Dashes = 1;
	    dash.DashLen = 150;
	    dash.Distance = 150;
	    lineXShapePropSet.setPropertyValue("LineDash", dash);
	} else {
	    lineXShapePropSet.setPropertyValue("LineStyle", LineStyle.SOLID);
	}

	return lineXShape;
    }

    public XShape drawEllipse(int x, int y, 
			     int width, int height) 
	throws java.lang.Exception {

	//System.out.println("UseCase("+x+", "+y+", "+width+", "+height+", "+name+")");
	x = (int)(dx * x);	
	y = (int)(dy * y);

	width = (int)(dx * width);
	height = (int)(dx * height);

	//System.out.println("  -->["+x+", "+y+", "+width+", "+height+"]");

	// 
	// Circle Object 
	//
	Object shape = xDocumentFactory.createInstance("com.sun.star.drawing.EllipseShape");
	XShape xShape = 
	    (XShape)UnoRuntime.queryInterface(XShape.class, shape);
	xShape.setPosition(new Point(x, y));
	xShape.setSize(new Size(width, height));
	
	// ucXShapes.add(xShape);
	xDrawPage.add(xShape);
	
	XPropertySet xShapePropSet = Soffice.getXPropertySet(xShape);
	
	// xShapePropSet.setPropertyValue("TextFitToSize", TextFitToSizeType.PROPORTIONAL );
	xShapePropSet.setPropertyValue("FillColor", new Integer(0xFFFFFF));
	// 	xShapePropSet.setPropertyValue("TextLeftDistance",  new Integer(500) );
	// 	xShapePropSet.setPropertyValue("TextRightDistance", new Integer(500) );
	// 	xShapePropSet.setPropertyValue("TextUpperDistance", new Integer(500) );
	// 	xShapePropSet.setPropertyValue("TextLowerDistance", new Integer(500) );

	return xShape;
    }


    public XShape addInheritanceTriangle(XShape modelObject) 
	throws java.lang.Exception {

	Object ucObj = msf.createInstance("com.sun.star.drawing.ShapeCollection");
	XShapes ucXShapes = (XShapes) UnoRuntime.queryInterface(XShapes.class, ucObj);

	ucXShapes.add(modelObject);

	Point moPosition = modelObject.getPosition();
	Size moSize = modelObject.getSize();
	
	// 
	// Polygon Inheritance Triangle
	//
	// 	XShape polyShape = 
	// 	    createShape(xDrawComponent,
	// 			new Point( 0, 0 ),
	// 			new Size( 0, 0 ),
	// 			"com.sun.star.drawing.ClosedBezierShape" );
	Object xObj = xDocumentFactory.createInstance( "com.sun.star.drawing.ClosedBezierShape" );
	XShape polyShape = (XShape)UnoRuntime.queryInterface(XShape.class, xObj );

	// polyShape.setPosition(new Point(6000, 8000));
	// polyShape.setSize(new Size(4000, 3000));  
	xDrawPage.add(polyShape);
	ucXShapes.add(polyShape);

	XPropertySet xShapeProperties = 
	    (XPropertySet)UnoRuntime.queryInterface(XPropertySet.class, 
						    polyShape);

	PolyPolygonBezierCoords aCoords = new PolyPolygonBezierCoords();
	// allocating the outer sequence
	aCoords.Coordinates = new Point[1][ ];	
	aCoords.Flags = new PolygonFlags[1][ ];

	int nPointCount = 4;
	Point[]		pPolyPoints     = new Point[ nPointCount ];
	PolygonFlags[]	pPolyFlags	= new PolygonFlags[ nPointCount ];
	
	for(int n = 0; n < nPointCount; n++ ) {
	    pPolyPoints[ n ] = new Point();
	}
	
	// xShape.setPosition(new Point(1000, 1000));
	// xShape.setSize(new Size(4000, 3000));  
	int x = moPosition.X + (moSize.Width / 2) - 250;
	int y = moPosition.Y + moSize.Height + 500;
	
	pPolyPoints[ 0 ].X = x + 0;
	pPolyPoints[ 0 ].Y = y + 0;
	pPolyFlags[ 0 ] = PolygonFlags.NORMAL;
	pPolyPoints[ 1 ].X = x + 250;
	pPolyPoints[ 1 ].Y = y - 500;
	pPolyFlags[ 1 ] = PolygonFlags.NORMAL;
	pPolyPoints[ 2 ].X = x + 500;
	pPolyPoints[ 2 ].Y = y + 0;
	pPolyFlags[ 2 ] = PolygonFlags.NORMAL;
	pPolyPoints[ 3 ].X = x + 0;
	pPolyPoints[ 3 ].Y = y + 0;
	pPolyFlags[ 3 ] = PolygonFlags.NORMAL;
	//
	aCoords.Coordinates[0] = pPolyPoints;
	aCoords.Flags[0]       = pPolyFlags;
	//
	xShapeProperties.setPropertyValue( "PolyPolygonBezier", aCoords );
	// White fill
	xShapeProperties.setPropertyValue("FillColor", new Integer(0xFFFFFF));
	// move the shape to the back by changing the ZOrder
	xShapeProperties.setPropertyValue( "ZOrder", new Integer( 1 ) );

	// Group UC and inheritance objects
	XShapeGrouper ucgr = (XShapeGrouper)UnoRuntime.queryInterface(XShapeGrouper.class, xDrawPage);
	XShape ucGroupXShape = ucgr.group(ucXShapes);

	return polyShape;
    }

    /**
     * Used to draw Inheritance Line
     */
    public XShape drawRectLine(XShape from, XShape to) 
	throws java.lang.Exception {
	//
	// Inhertiance Connector
	//
	Object connector = xDocumentFactory. createInstance("com.sun.star.drawing.ConnectorShape") ;
	com.sun.star.drawing.XShape xConnector = 
	    (com.sun.star.drawing.XShape)UnoRuntime.queryInterface(com.sun.star.drawing.XShape.class, 
								   connector);                    
	xDrawPage.add(xConnector);
	com.sun.star.beans.XPropertySet xConnectorProps = 
	    (com.sun.star.beans.XPropertySet)UnoRuntime.queryInterface(com.sun.star.beans.XPropertySet.class, 
								       connector);

	xConnectorProps.setPropertyValue("StartShape", to);
	xConnectorProps.setPropertyValue("EndShape", from);

	// glue point positions: 0=top 1=left 2=bottom 3=right
	xConnectorProps.setPropertyValue("StartGluePointIndex", new Integer(2));
	xConnectorProps.setPropertyValue("EndGluePointIndex", new Integer(0)); 
	
	return xConnector;
    }


    /**
     * @return the XShape of the whole Group of objects
     */
    public XShape drawActor(int x, int y, 
			    int width, int height,
			    String name) 
	throws java.lang.Exception {

	x = (int)(dx * x);	
	y = (int)(dy * y);

	// Head Size
	width = (int)(dx * width);
	height = (int)(dy * height);
	int head_height = (int)(height / 3.0f);
	int arm_height = 2 * (int)(height / 5.0f);
	int body_height = (int)(height / 3.0f);
	int legs_height = (int)(height / 3.0f);
		
	// 
	// Actor Object 
	//
	Object shObj = msf.createInstance("com.sun.star.drawing.ShapeCollection");
	XShapes actorXShapes = (XShapes) UnoRuntime.queryInterface(XShapes.class, shObj);

	Object actorShape = xDocumentFactory.createInstance("com.sun.star.drawing.EllipseShape");
	XShape actorXShape = 
	    (XShape)UnoRuntime.queryInterface(XShape.class, actorShape);
	actorXShape.setPosition(new Point(x, y));
	actorXShape.setSize(new Size(width, head_height));  
	xDrawPage.add(actorXShape);
	actorXShapes.add(actorXShape);
	XPropertySet actorxShapePropSet = Soffice.getXPropertySet(actorXShape);
	actorxShapePropSet.setPropertyValue("TextFitToSize", TextFitToSizeType.PROPORTIONAL );
	actorxShapePropSet.setPropertyValue("FillColor", new Integer(0xFFFFFF));
	actorxShapePropSet.setPropertyValue("TextLeftDistance",  new Integer(100) );
	actorxShapePropSet.setPropertyValue("TextRightDistance", new Integer(100) );
	actorxShapePropSet.setPropertyValue("TextUpperDistance", new Integer(100) );
	actorxShapePropSet.setPropertyValue("TextLowerDistance", new Integer(100) );

	// Draw Body Line
	Object actorBodyShape = xDocumentFactory.createInstance("com.sun.star.drawing.LineShape");
	XShape actorBodyShapeXShape = 
	    (XShape)UnoRuntime.queryInterface(XShape.class, actorBodyShape);
	actorBodyShapeXShape.setPosition(new Point(x + (int)(width / 2.0f), 
						   y + head_height));
	actorBodyShapeXShape.setSize(new Size(0, body_height));  
	xDrawPage.add(actorBodyShapeXShape);
	actorXShapes.add(actorBodyShapeXShape);
	XPropertySet actorBodyShapeXShapePropSet = Soffice.getXPropertySet(actorBodyShapeXShape);
	actorBodyShapeXShapePropSet.setPropertyValue("LineColor", new Integer(0));
	actorBodyShapeXShapePropSet.setPropertyValue("LineStyle", LineStyle.SOLID);

 	XPropertySet actorTextPropSet = Soffice.addPortion(actorBodyShapeXShape, name, false );
 	actorTextPropSet.setPropertyValue("ParaAdjust", ParagraphAdjust.CENTER );
 	actorTextPropSet.setPropertyValue("CharColor",  new Integer( SofficeDialog.COLOR_SUN1 ) );
  	actorBodyShapeXShapePropSet.setPropertyValue("TextHorizontalAdjust",  TextHorizontalAdjust.CENTER );
  	actorBodyShapeXShapePropSet.setPropertyValue("TextVerticalAdjust",  TextVerticalAdjust.BOTTOM );
 	actorBodyShapeXShapePropSet.setPropertyValue("TextLowerDistance", new Integer(-1000));
 	actorBodyShapeXShapePropSet.setPropertyValue("CharFontName", "Arial");
 	actorBodyShapeXShapePropSet.setPropertyValue("CharHeight", new Float(10));

	// Arms
	Object actorArmsShape = xDocumentFactory.createInstance("com.sun.star.drawing.LineShape");
	XShape actorArmsShapeXShape = 
	    (XShape)UnoRuntime.queryInterface(XShape.class, actorArmsShape);
	// actorArmsShapeXShape.setPosition(new Point(x + 250 - 300, y + 500 + 150));
	actorArmsShapeXShape.setPosition(new Point(x, 
						   y + arm_height));
	actorArmsShapeXShape.setSize(new Size(width, 
					      0));  
	xDrawPage.add(actorArmsShapeXShape);
	actorXShapes.add(actorArmsShapeXShape);
	XPropertySet actorArmsShapeXShapePropSet = Soffice.getXPropertySet(actorArmsShapeXShape);
	actorArmsShapeXShapePropSet.setPropertyValue("LineColor", new Integer(0));
	actorArmsShapeXShapePropSet.setPropertyValue("LineStyle", LineStyle.SOLID);

	// Left Leg
	Object actorLeftLegShape = xDocumentFactory.createInstance("com.sun.star.drawing.LineShape");
	XShape actorLeftLegShapeXShape = 
	    (XShape)UnoRuntime.queryInterface(XShape.class, actorLeftLegShape);
	actorLeftLegShapeXShape.setPosition(new Point(x + (int)(width / 2.0f), 
						      y + head_height + body_height));
	actorLeftLegShapeXShape.setSize(new Size(0 - (int)(width / 2.0f), 
						 legs_height));  
	xDrawPage.add(actorLeftLegShapeXShape);
	actorXShapes.add(actorLeftLegShapeXShape);
	XPropertySet actorLeftLegShapeXShapePropSet = Soffice.getXPropertySet(actorLeftLegShapeXShape);
	actorLeftLegShapeXShapePropSet.setPropertyValue("LineColor", new Integer(0));
	actorLeftLegShapeXShapePropSet.setPropertyValue("LineStyle", LineStyle.SOLID);

	// Right Leg
	Object actorRightLegShape = xDocumentFactory.createInstance("com.sun.star.drawing.LineShape");
	XShape actorRightLegShapeXShape = 
	    (XShape)UnoRuntime.queryInterface(XShape.class, actorRightLegShape);
	// actorRightLegShapeXShape.setPosition(new Point(x + 250, y + 500 + 500));
	actorRightLegShapeXShape.setPosition(new Point(x + (int)(width / 2.0f), 
						       y + head_height + body_height));
	actorRightLegShapeXShape.setSize(new Size((int)(width / 2.0f), 
						  legs_height));  
	xDrawPage.add(actorRightLegShapeXShape);
	actorXShapes.add(actorRightLegShapeXShape);
	XPropertySet actorRightLegShapeXShapePropSet = Soffice.getXPropertySet(actorRightLegShapeXShape);
	actorRightLegShapeXShapePropSet.setPropertyValue("LineColor", new Integer(0));
	actorRightLegShapeXShapePropSet.setPropertyValue("LineStyle", LineStyle.SOLID);

	XShapeGrouper gr = (XShapeGrouper)UnoRuntime.queryInterface(XShapeGrouper.class, xDrawPage);
	XShape actorGroupXShape = gr.group(actorXShapes);
	// xDrawPage.add(actorGroupXShape);

	return actorGroupXShape;
    }


    public XShape drawClass(int x, int y, 
			    int width, int height,
			    String name) 
	throws java.lang.Exception {

	
	x = (int)(dx * x);	
	y = (int)(dy * y);

	width = (int)(dx * width);
	height = (int)(dy * height);
	
	//
	// Class Rectangle
	// 
	Object shape2 = xDocumentFactory.createInstance("com.sun.star.drawing.RectangleShape");
	XShape xShape2 = 
	    (XShape)UnoRuntime.queryInterface(XShape.class, shape2);
	xShape2.setPosition(new Point(x, y));
	xShape2.setSize(new Size(width, height));  

	xDrawPage.add(xShape2);

	XPropertySet xShapePropSet2 = Soffice.getXPropertySet(xShape2);
	
	//xShapePropSet2.setPropertyValue("TextFitToSize", TextFitToSizeType.PROPORTIONAL );
	xShapePropSet2.setPropertyValue("FillColor", new Integer(0xFFFFFF));
	// 	xShapePropSet2.setPropertyValue("TextLeftDistance",  new Integer(500) );
	// 	xShapePropSet2.setPropertyValue("TextRightDistance", new Integer(500) );
	// 	xShapePropSet2.setPropertyValue("TextUpperDistance", new Integer(500) );
	// 	xShapePropSet2.setPropertyValue("TextLowerDistance", new Integer(500) );

	XPropertySet xTextPropSet2 = Soffice.addPortion(xShape2, name, false );
	xTextPropSet2.setPropertyValue("ParaAdjust", ParagraphAdjust.CENTER );
	//xTextPropSet2.setPropertyValue("CharColor",  new Integer( SofficeDialog.COLOR_SUN1 ) );
	xTextPropSet2.setPropertyValue("CharColor", new Integer(0));
	xTextPropSet2.setPropertyValue("CharHeight", new Integer(10));
 	xTextPropSet2.setPropertyValue("CharFontName", "Arial");

	// 	xTextPropSet2 = Soffice.addPortion(xShape2, "row-2", true );
	// 	xTextPropSet2.setPropertyValue("CharColor",  new Integer( SofficeDialog.COLOR_SUN2 ) );
	//  	xTextPropSet2.setPropertyValue("CharFontName", "Arial");

	return xShape2;
    }

    public XShape drawComponent(int x, int y, 
				int width, int height,
				String name) 
	throws java.lang.Exception {

	
	x = (int)(dx * x);	
	y = (int)(dy * y);

	width = (int)(dx * width);
	height = (int)(dy * height);
	
	Object compObj = msf.createInstance("com.sun.star.drawing.ShapeCollection");
	XShapes compXShapes = (XShapes) UnoRuntime.queryInterface(XShapes.class, compObj);

	//
	// Main Rect
	//
	Object mainRect = xDocumentFactory.createInstance("com.sun.star.drawing.RectangleShape");
	XShape xMainRect = 
	    (XShape)UnoRuntime.queryInterface(XShape.class, mainRect);
	xMainRect.setPosition(new Point(x, y));
	xMainRect.setSize(new Size(width, height));  
	xDrawPage.add(xMainRect);
	compXShapes.add(xMainRect);
	XPropertySet mainRectProps = Soffice.getXPropertySet(xMainRect);
	mainRectProps.setPropertyValue("FillColor", new Integer(0xFFFFFF));
	XPropertySet mainRectTextProps = Soffice.addPortion(xMainRect, name, false );
	mainRectTextProps.setPropertyValue("ParaAdjust", ParagraphAdjust.CENTER );
	mainRectTextProps.setPropertyValue("CharColor", new Integer(0));
	mainRectTextProps.setPropertyValue("CharHeight", new Integer(10));
 	mainRectTextProps.setPropertyValue("CharFontName", "Arial");

	int box_width = 600;
	int box_height = 300;

	//
	// Top Box
	//
	Object topRect = xDocumentFactory.createInstance("com.sun.star.drawing.RectangleShape");
	XShape xTopRect = 
	    (XShape)UnoRuntime.queryInterface(XShape.class, topRect);
	xTopRect.setPosition(new Point(x - 200, y + 200));
	xTopRect.setSize(new Size(box_width, box_height));  
	xDrawPage.add(xTopRect);
	compXShapes.add(xTopRect);
	XPropertySet topRectProps = Soffice.getXPropertySet(xTopRect);
	topRectProps.setPropertyValue("FillColor", new Integer(0xFFFFFF));

	//
	// Bottom Box
	//
	Object bottomrect = xDocumentFactory.createInstance("com.sun.star.drawing.RectangleShape");
	XShape xBottomrect = 
	    (XShape)UnoRuntime.queryInterface(XShape.class, bottomrect);
	xBottomrect.setPosition(new Point(x - 200, y + box_height + 200 + 200));
	xBottomrect.setSize(new Size(box_width, box_height));  
	xDrawPage.add(xBottomrect);
	compXShapes.add(xBottomrect);
	XPropertySet bottomrectProps = Soffice.getXPropertySet(xBottomrect);
	bottomrectProps.setPropertyValue("FillColor", new Integer(0xFFFFFF));

	XShapeGrouper gr = (XShapeGrouper)UnoRuntime.queryInterface(XShapeGrouper.class, xDrawPage);
	XShape compGroupXShape = gr.group(compXShapes);

	return compGroupXShape;
    }

    public XShape drawAssociation(XShape from, 
				   XShape to) 
	throws java.lang.Exception {
	return drawAssociation(from, to, false, false, null, null, null);
    }

    public XShape drawAssociation(XShape from, 
				   XShape to,
				   boolean fromArrow,
				   boolean toArrow) 
	throws java.lang.Exception {
	return drawAssociation(from, to, fromArrow, toArrow, null, null, null);
    }

    public XShape drawAssociation(XShape from, 
				   XShape to,
				   boolean fromArrow,
				   boolean toArrow,
				   String text,
				   String fromText,
				   String toText) 
	throws java.lang.Exception {
	//
	// Association Connector
	//
	Object assConnector = xDocumentFactory. createInstance("com.sun.star.drawing.ConnectorShape") ;
	com.sun.star.drawing.XShape assxConnector = 
	    (com.sun.star.drawing.XShape)UnoRuntime.queryInterface(com.sun.star.drawing.XShape.class, 
								   assConnector);                    
	xDrawPage.add(assxConnector);
	com.sun.star.beans.XPropertySet assxConnectorProps = 
	    (com.sun.star.beans.XPropertySet)UnoRuntime.queryInterface(com.sun.star.beans.XPropertySet.class, 
								       assConnector);

	assxConnectorProps.setPropertyValue("StartShape", from);
	assxConnectorProps.setPropertyValue("EndShape", to);

 	assxConnectorProps.setPropertyValue("EdgeKind", ConnectorType.LINE);
	if(fromArrow) {
	    assxConnectorProps.setPropertyValue("LineStartName", "Line Arrow");
	}
	if(toArrow) {
	    assxConnectorProps.setPropertyValue("LineEndName", "Line Arrow");
	}

	Point lineStart = 
	    (Point)assxConnectorProps.getPropertyValue("StartPosition");
	Point lineEnd = 
	    (Point)assxConnectorProps.getPropertyValue("EndPosition");

	int dx = lineEnd.X - lineStart.X;
	int dy = lineEnd.Y - lineStart.Y;
	double startAngle = Math.atan2((double)dy, (double)dx);
	double startAngleDeg = Math.toDegrees(startAngle);
	double endAngle = Math.atan2((double)dx, (double)dy);
	double endAngleDeg = Math.toDegrees(endAngle);

// 	System.out.println("startAngle="+startAngleDeg+
// 			   ", lineStart={"+lineStart.X+", "+lineStart.Y+"}"+
// 			   ", lineEnd={"+lineEnd.X+", "+lineEnd.Y+"}");

	int start_dx = 0;
	int start_dy = 0;
	if(startAngleDeg > 0.0 && startAngleDeg <= 45.0) {
	    start_dx = 1000;
	    start_dy = -1000;
	} else if(startAngleDeg > 45.0 && startAngleDeg <= 90.0) {
	    start_dx = 1000;
	    start_dy = -1000;
	} else if(startAngleDeg > 90.0 && startAngleDeg <= 135.0) {
	    start_dx = -1000;
	    start_dy = -1000;
	} else if(startAngleDeg > 135.0) {
	    start_dx = -1000;
	    start_dy = -1000;
	} else if(startAngleDeg < 0.0 && startAngleDeg >= -45.0) {
	    start_dx = 1000;
	    start_dy = 1000;	    
	} else if(startAngleDeg < -45.0 && startAngleDeg >= -90.0) {
	    start_dx = 1000;
	    start_dy = 1000;	    
	} else if(startAngleDeg < -90.0 && startAngleDeg >= -135.0) {
	    start_dx = -1000;
	    start_dy = 1000;	    
	} else if(startAngleDeg < -135.0) {
	    start_dx = -1000;
	    start_dy = 1000;	    
	}

	int end_dx = 0;
	int end_dy = 0;
	if(endAngleDeg > 0.0 && endAngleDeg <= 45.0) {
	    end_dx = 1000;
	    end_dy = -1000;
	} else if(endAngleDeg > 45.0 && endAngleDeg <= 90.0) {
	    end_dx = 1000;
	    end_dy = -1000;
	} else if(endAngleDeg > 90.0 && endAngleDeg <= 135.0) {
	    end_dx = -1000;
	    end_dy = -1000;
	} else if(endAngleDeg > 135.0) {
	    end_dx = -1000;
	    end_dy = -1000;
	} else if(endAngleDeg < 0.0 && endAngleDeg >= -45.0) {
	    end_dx = 1000;
	    end_dy = 1000;	    
	} else if(endAngleDeg < -45.0 && endAngleDeg >= -90.0) {
	    end_dx = 1000;
	    end_dy = 1000;	    
	} else if(endAngleDeg < -90.0 && endAngleDeg >= -135.0) {
	    end_dx = -1000;
	    end_dy = 1000;	    
	} else if(endAngleDeg < -135.0) {
	    end_dx = -1000;
	    end_dy = 1000;	    
	}

	

	// Add Text to line
	if(text != null) {
	    XPropertySet textProps = Soffice.addPortion(assxConnector, text, false );
	    textProps.setPropertyValue("CharColor",  new Integer(0x000000) );
	    textProps.setPropertyValue("CharHeight",  new Integer(10) );
	    textProps.setPropertyValue("CharFontName", "Arial");	
	    textProps.setPropertyValue("ParaBottomMargin", new Integer(300));
	}

	if(fromText != null) {
	    //
	    // Assocication Start Text
	    // 
	    Object startShape = xDocumentFactory.createInstance("com.sun.star.drawing.RectangleShape");
	    XShape xStartShape = 
		(XShape)UnoRuntime.queryInterface(XShape.class, startShape);
	    xStartShape.setPosition(new Point(lineStart.X + start_dx, lineStart.Y - start_dy));
	    // xStartShape.setSize(new Size(width, height));  
	    xDrawPage.add(xStartShape);
	    XPropertySet startShapeProps = Soffice.getXPropertySet(xStartShape);
	    startShapeProps.setPropertyValue("FillColor", new Integer(0xFFFFFF));
	    startShapeProps.setPropertyValue("LineColor", new Integer(0xFFFFFF));
	    XPropertySet startTextProps = Soffice.addPortion(xStartShape, fromText, false );
	    // startShapeProps.setPropertyValue("TextAutoGrowWidth", new Boolean(true));
	    // startShapeProps.setPropertyValue("TextAutoGrowHeight", new Boolean(true));
	    startTextProps.setPropertyValue("CharColor",  new Integer(0x000000) );
	    startTextProps.setPropertyValue("CharHeight",  new Integer(10) );
	    startTextProps.setPropertyValue("CharFontName", "Arial");	
	}

	if(toText != null) {
	    //
	    // Assocication End Text
	    // 
	    Object endShape = xDocumentFactory.createInstance("com.sun.star.drawing.RectangleShape");
	    XShape xEndShape = 
		(XShape)UnoRuntime.queryInterface(XShape.class, endShape);
	    xEndShape.setPosition(new Point(lineEnd.X + end_dx, lineEnd.Y - end_dy));
	    // xEndShape.setSize(new Size(width, height));  
	    xDrawPage.add(xEndShape);
	    XPropertySet endShapeProps = Soffice.getXPropertySet(xEndShape);
	    endShapeProps.setPropertyValue("FillColor", new Integer(0xFFFFFF));
	    endShapeProps.setPropertyValue("LineColor", new Integer(0xFFFFFF));
	    XPropertySet endTextProps = Soffice.addPortion(xEndShape, toText, false );
	    endTextProps.setPropertyValue("CharColor",  new Integer(0x000000) );
	    endTextProps.setPropertyValue("CharHeight",  new Integer(10) );
	    endTextProps.setPropertyValue("CharFontName", "Arial");	
	}

	return assxConnector;
    }

    public void addNewDrawPage(String name) 
	throws java.lang.Exception {

	XDrawPage xLastPage = addNewDrawPage();
	
	XNamed xNamed = (XNamed)UnoRuntime.queryInterface(XNamed.class, 
							  xLastPage );	
	// beware, the page must have an unique name
	xNamed.setName( name );
    }

    /**
     * Set name of last page
     */
    public void setDrawPageName(String name) 
	throws java.lang.Exception {

	int nDrawPages = getDrawPageCount(xDrawComponent);	
	XDrawPage xLastPage = getDrawPageByIndex(xDrawComponent, nDrawPages - 1 );
	
	XNamed xNamed = (XNamed)UnoRuntime.queryInterface(XNamed.class, 
							  xLastPage );
	
	// beware, the page must have an unique name
	xNamed.setName( name );

    }
    

    public XDrawPage addNewDrawPage() 
	throws java.lang.Exception {

	int nDrawPages = getDrawPageCount(xDrawComponent);

	xDrawPage = insertNewDrawPageByIndex( xDrawComponent, nDrawPages );	
	return xDrawPage;
    }

    static public XDrawPage insertNewDrawPageByIndex(XComponent xComponent, 
						     int nIndex )
	throws java.lang.Exception {
	XDrawPagesSupplier xDrawPagesSupplier =
	    (XDrawPagesSupplier)UnoRuntime.queryInterface(XDrawPagesSupplier.class, 
							  xComponent );
	XDrawPages xDrawPages = xDrawPagesSupplier.getDrawPages();
	return xDrawPages.insertNewByIndex( nIndex );
    }

    static public int getDrawPageCount( XComponent xComponent )
    {
	XDrawPagesSupplier xDrawPagesSupplier= 
	    (XDrawPagesSupplier)UnoRuntime.queryInterface(XDrawPagesSupplier.class, xComponent );
	XDrawPages xDrawPages = xDrawPagesSupplier.getDrawPages();
	return xDrawPages.getCount();
    }

    static public XDrawPage getDrawPageByIndex( XComponent xComponent, int nIndex )
	throws com.sun.star.lang.IndexOutOfBoundsException,
	       com.sun.star.lang.WrappedTargetException {

	XDrawPagesSupplier xDrawPagesSupplier =
	    (XDrawPagesSupplier)UnoRuntime.queryInterface(XDrawPagesSupplier.class, xComponent );
	XDrawPages xDrawPages = xDrawPagesSupplier.getDrawPages();
	return (XDrawPage)UnoRuntime.queryInterface(XDrawPage.class, xDrawPages.getByIndex( nIndex ));
    }
    
}

