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
import java.util.Vector;

import org.apache.log4j.Logger;

import com.sun.star.awt.Point;
import com.sun.star.awt.Size;
import com.sun.star.beans.XPropertySet;
import com.sun.star.drawing.ConnectorType;
import com.sun.star.drawing.LineStyle;
import com.sun.star.drawing.PolyPolygonBezierCoords;
import com.sun.star.drawing.PolygonFlags;
import com.sun.star.drawing.TextFitToSizeType;
import com.sun.star.drawing.TextHorizontalAdjust;
import com.sun.star.drawing.TextVerticalAdjust;
import com.sun.star.drawing.XShape;
import com.sun.star.drawing.XShapes;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.style.ParagraphAdjust;
import com.sun.star.text.TextContentAnchorType;
import com.sun.star.text.WrapTextMode;
import com.sun.star.text.XText;
import com.sun.star.text.XTextContent;
import com.sun.star.uno.UnoRuntime;

/**
 * Class DrawCache. This classes produces objects which are later instantiated as
 * StarOffice objects when constructing the document.
 * 
 * @author Matti Pehrs
 *
 */
public class DrawCache {

    private static Logger LOG = Logger.getLogger(DrawCache.class);
    
    /**
     * This class is used to cache a Shape.
     */
    private static class DrawCacheShape {
        public DrawCacheShape(java.lang.Object shapeObj, XShape xshape) {
            this.shapeObj = shapeObj;
            this.xshape = xshape;
        }

        public java.lang.Object shapeObj;

        public XShape xshape;

        public Vector properties = new Vector(); // [DrawCacheProperty]

        public String text = null;

        public Vector textProperties = new Vector(); // [DrawCacheProperty]
    };

    /**
     * This class is used to cache properties.
     */
    private static class DrawCacheProperty {
        public DrawCacheProperty(String name, java.lang.Object value) {
            this.name = name;
            this.value = value;
        }

        public String name;

        public java.lang.Object value;
    };

    /**
     * The cache of Shapes.
     */
    private Vector cache = new Vector(); // [DrawCacheShape]

    /**
     * Search a shape object in the cache.
     * @param shapeObj
     * @return
     */
    private DrawCacheShape findDrawCacheShape(java.lang.Object shapeObj) {
        for (Iterator it = cache.iterator(); it.hasNext();) {
            DrawCacheShape shape = (DrawCacheShape) it.next();
            if (shape.shapeObj == shapeObj) {
                return shape;
            }
        }
        return null;
    }

    /**
     * Add a shape to the cache.
     * @param shapeObj
     * @param xshape
     */
    public void addShape(java.lang.Object shapeObj, XShape xshape) {
        this.cache.add(new DrawCacheShape(shapeObj, xshape));
    }

    /**
     * Add a property to a (cached) shape.
     * @param shapeObj
     * @param name
     * @param value
     * @throws java.lang.Exception
     */
    public void setShapeProperty(java.lang.Object shapeObj, String name,
            java.lang.Object value) throws java.lang.Exception {

        DrawCacheShape shape = findDrawCacheShape(shapeObj);
        if (shape == null) {
            throw new java.lang.Exception("Cannot find cached Shape");
        }
        shape.properties.add(new DrawCacheProperty(name, value));
    }

    /**
     * Set the text for a (cached) shape.
     * @param shapeObj
     * @param text
     * @throws java.lang.Exception
     */
    public void setShapeText(java.lang.Object shapeObj, String text)
            throws java.lang.Exception {

        DrawCacheShape shape = findDrawCacheShape(shapeObj);
        if (shape == null) {
            throw new java.lang.Exception("Cannot find cached Shape");
        }
        shape.text = text;
    }

    /**
     * Set a text property on the (cached) shape.
     * @param shapeObj
     * @param name
     * @param value
     * @throws java.lang.Exception
     */
    public void setShapeTextProperty(java.lang.Object shapeObj, String name,
            java.lang.Object value) throws java.lang.Exception {

        DrawCacheShape shape = findDrawCacheShape(shapeObj);
        if (shape == null) {
            throw new java.lang.Exception("Cannot find cached Shape");
        }
        shape.textProperties.add(new DrawCacheProperty(name, value));
    }

    /**
     * 
     * @param msf
     * @param text
     * @throws java.lang.Exception
     */
    public void addCache2Text(XMultiServiceFactory msf, XText text)
            throws java.lang.Exception {

        //
        // Create and add group
        //
        Object mgroupobj = msf
                .createInstance("com.sun.star.drawing.GroupShape");
        XShape mgroup = (XShape) UnoRuntime.queryInterface(XShape.class,
                mgroupobj);

        // Set Group Properties
        XPropertySet mgroupProps = (XPropertySet) UnoRuntime.queryInterface(
                XPropertySet.class, mgroupobj);
        mgroupProps.setPropertyValue("AnchorType",
                TextContentAnchorType.AS_CHARACTER);
        mgroupProps.setPropertyValue("TextWrap", WrapTextMode.NONE);

        // Add Group Shape to text doc
        // query for the shape collection of xDrawPage
        XTextContent mgroupTextContent = (XTextContent) UnoRuntime
                .queryInterface(XTextContent.class, mgroupobj);
        text.insertTextContent(text.getEnd(), mgroupTextContent, false);
        XShapes mgroupShapes = (XShapes) UnoRuntime.queryInterface(
                XShapes.class, mgroupobj);

        mgroupProps = (XPropertySet) UnoRuntime.queryInterface(
                XPropertySet.class, mgroupobj);
        mgroupProps.setPropertyValue("AnchorType",
                TextContentAnchorType.AS_CHARACTER);
        mgroupProps.setPropertyValue("TextWrap", WrapTextMode.NONE);

        // Add shapes to the group
        for (Iterator it = cache.iterator(); it.hasNext();) {
            DrawCacheShape shape = (DrawCacheShape) it.next();
            mgroupShapes.add(shape.xshape);
        }

        // Set properties on shapes
        for (Iterator it = cache.iterator(); it.hasNext();) {
            DrawCacheShape shape = (DrawCacheShape) it.next();
            // Set BG Properties
            XPropertySet shapeProps = (XPropertySet) UnoRuntime.queryInterface(
                    XPropertySet.class, shape.shapeObj);
            for (Iterator pit = shape.properties.iterator(); pit.hasNext();) {
                DrawCacheProperty prop = (DrawCacheProperty) pit.next();
                shapeProps.setPropertyValue(prop.name, prop.value);
            }

            if (shape.text != null) {
                XText shText = (XText) UnoRuntime.queryInterface(XText.class,
                        shape.shapeObj);
                shText.insertString(shText.getEnd(), shape.text, false);

                XPropertySet shapeTextProps = (XPropertySet) UnoRuntime
                        .queryInterface(XPropertySet.class, shText);
                for (Iterator pit = shape.textProperties.iterator(); pit
                        .hasNext();) {
                    DrawCacheProperty prop = (DrawCacheProperty) pit.next();
                    shapeTextProps.setPropertyValue(prop.name, prop.value);
                }

            }

        }
    }

    public XShape addAssociationXShape(XMultiServiceFactory xDocumentFactory,
            XShape from, XShape to) throws java.lang.Exception {
        //
        // Association Connector
        //
        Object assConnector = xDocumentFactory
                .createInstance("com.sun.star.drawing.ConnectorShape");
        com.sun.star.drawing.XShape assxConnector = (com.sun.star.drawing.XShape) UnoRuntime
                .queryInterface(com.sun.star.drawing.XShape.class, assConnector);
        addShape(assConnector, assxConnector);

        setShapeProperty(assConnector, "StartShape", from);
        setShapeProperty(assConnector, "EndShape", to);

        setShapeProperty(assConnector, "EdgeKind", ConnectorType.LINE);
        setShapeProperty(assConnector, "LineEndName", "Line Arrow");

        return assxConnector;
    }

    /**
     * @return the XShape of the whole Group of objects
     */
    public XShape addActorXShape(XMultiServiceFactory xDocumentFactory, int x,
            int y, int width, int height, String name)
            throws java.lang.Exception {

        LOG.info("Draw actor '"+name+"',x="+x+",width="+width+",y="+y+",height="+height);
        // x = (int)(dx * x);
        // y = (int)(dy * y);
        // // Head Size
        // width = (int)(dx * width);
        // height = (int)(dy * height);

        int head_height = (int) (height / 3.0f);
        int arm_height = 2 * (int) (height / 5.0f);
        int body_height = (int) (height / 3.0f);
        int legs_height = (int) (height / 3.0f);

        // 
        // Actor Object
        //
        // Object mgroupobj =
        // xDocumentFactory.createInstance("com.sun.star.drawing.GroupShape");
        // XShape mgroup = (XShape)UnoRuntime.queryInterface(XShape.class,
        // mgroupobj);
        // System.out.println("mgroupobj="+mgroupobj);
        // System.out.println("mgroup="+mgroup);
        // // addShape(mgroupobj, mgroup);
        // XShapes actorXShapes =
        // (XShapes)UnoRuntime.queryInterface(XShapes.class, mgroupobj);

        // Actor Head
        Object actorShape = xDocumentFactory
                .createInstance("com.sun.star.drawing.EllipseShape");
        XShape actorXShape = (XShape) UnoRuntime.queryInterface(XShape.class,
                actorShape);
        actorXShape.setPosition(new Point(x, y));
        actorXShape.setSize(new Size(width, head_height));
        addShape(actorShape, actorXShape);
        // actorXShapes.add(actorXShape);
        // Props
        setShapeProperty(actorShape, "TextFitToSize",
                TextFitToSizeType.PROPORTIONAL);
        setShapeProperty(actorShape, "FillColor", new Integer(0xFFFFFF));
        setShapeProperty(actorShape, "TextLeftDistance", new Integer(100));
        setShapeProperty(actorShape, "TextRightDistance", new Integer(100));
        setShapeProperty(actorShape, "TextUpperDistance", new Integer(100));
        setShapeProperty(actorShape, "TextLowerDistance", new Integer(100));

        // Draw Body Line
        Object actorBodyShape = xDocumentFactory
                .createInstance("com.sun.star.drawing.LineShape");
        XShape actorBodyShapeXShape = (XShape) UnoRuntime.queryInterface(
                XShape.class, actorBodyShape);
        actorBodyShapeXShape.setPosition(new Point(x + (int) (width / 2.0f), y
                + head_height));
        actorBodyShapeXShape.setSize(new Size(0, body_height));
        addShape(actorBodyShape, actorBodyShapeXShape);
        // actorXShapes.add(actorBodyShapeXShape);
        // Props
        setShapeProperty(actorBodyShape, "LineColor", new Integer(0));
        setShapeProperty(actorBodyShape, "LineStyle", LineStyle.SOLID);
        setShapeProperty(actorBodyShape, "TextHorizontalAdjust",
                TextHorizontalAdjust.CENTER);
        setShapeProperty(actorBodyShape, "TextVerticalAdjust",
                TextVerticalAdjust.BOTTOM);
        setShapeProperty(actorBodyShape, "TextLowerDistance",
                new Integer(-1000));
        setShapeProperty(actorBodyShape, "CharFontName", "Arial");
        setShapeProperty(actorBodyShape, "CharHeight", new Float(10));

        // FIXME: We need to add the ability to set the properties on the text
        // as well as on the XShape
        // XPropertySet actorTextPropSet =
        // Soffice.addPortion(actorBodyShapeXShape, name, false );
        // actorTextPropSet.setPropertyValue("ParaAdjust",
        // ParagraphAdjust.CENTER );
        // actorTextPropSet.setPropertyValue("CharColor", new Integer(
        // SofficeDialog.COLOR_SUN1 ) );
        setShapeText(actorBodyShape, name);
        setShapeTextProperty(actorBodyShape, "ParaAdjust",
                ParagraphAdjust.CENTER);
        setShapeTextProperty(actorBodyShape, "CharColor", new Integer(
                SofficeDialog.COLOR_SUN1));
        setShapeTextProperty(actorBodyShape, "CharFontName", "Arial");

        // Arms
        Object actorArmsShape = xDocumentFactory
                .createInstance("com.sun.star.drawing.LineShape");
        XShape actorArmsShapeXShape = (XShape) UnoRuntime.queryInterface(
                XShape.class, actorArmsShape);
        actorArmsShapeXShape.setPosition(new Point(x, y + arm_height));
        actorArmsShapeXShape.setSize(new Size(width, 0));
        addShape(actorArmsShape, actorArmsShapeXShape);
        // actorXShapes.add(actorArmsShapeXShape);
        setShapeProperty(actorArmsShape, "LineColor", new Integer(0));
        setShapeProperty(actorArmsShape, "LineStyle", LineStyle.SOLID);

        // Left Leg
        Object actorLeftLegShape = xDocumentFactory
                .createInstance("com.sun.star.drawing.LineShape");
        XShape actorLeftLegShapeXShape = (XShape) UnoRuntime.queryInterface(
                XShape.class, actorLeftLegShape);
        actorLeftLegShapeXShape.setPosition(
                new Point(x + (int) (width / 2.0f),
                y + head_height + body_height));
        actorLeftLegShapeXShape.setSize(
                new Size(0 - (int) (width / 2.0f),
                legs_height));
        addShape(actorLeftLegShape, actorLeftLegShapeXShape);
        // actorXShapes.add(actorLeftLegShapeXShape);
        setShapeProperty(actorLeftLegShape, "LineColor", new Integer(0));
        setShapeProperty(actorLeftLegShape, "LineStyle", LineStyle.SOLID);

        // Right Leg
        Object actorRightLegShape = xDocumentFactory
                .createInstance("com.sun.star.drawing.LineShape");
        XShape actorRightLegShapeXShape = (XShape) UnoRuntime.queryInterface(
                XShape.class, actorRightLegShape);
        // actorRightLegShapeXShape.setPosition(new Point(x + 250, y + 500 +
        // 500));
        actorRightLegShapeXShape.setPosition(
                new Point( x + (int) (width / 2.0f),
                y + head_height + body_height));
        actorRightLegShapeXShape.setSize(
                new Size((int) (width / 2.0f),
                legs_height));
        addShape(actorRightLegShape, actorRightLegShapeXShape);
        // actorXShapes.add(actorRightLegShapeXShape);
        setShapeProperty(actorRightLegShape, "LineColor", new Integer(0));
        setShapeProperty(actorRightLegShape, "LineStyle", LineStyle.SOLID);

        // XShapeGrouper gr =
        // (XShapeGrouper)UnoRuntime.queryInterface(XShapeGrouper.class,
        // textDoc);
        // XShape actorGroupXShape = gr.group(actorXShapes);
        // xDrawPage.add(actorGroupXShape);

        return actorXShape;
        // return mgroup;
    }

    public XShape addInheritanceTriangleXShape(
            XMultiServiceFactory xDocumentFactory, XShape modelObject)
            throws java.lang.Exception {

        System.out.println("before createInstance...");

        // Object ucObj =
        // xDocumentFactory.createInstance("com.sun.star.drawing.ShapeCollection");
        // XShapes ucXShapes = (XShapes)
        // UnoRuntime.queryInterface(XShapes.class, ucObj);
        // System.out.println("ucObj="+ucObj);
        // System.out.println("ucXShapes="+ucXShapes);
        // ucXShapes.add(modelObject);

        Point moPosition = modelObject.getPosition();
        Size moSize = modelObject.getSize();

        System.out.println("modelObject.getPosition()=" + moPosition.X + ", "
                + moPosition.Y);
        System.out.println("modelObject.getSize()=" + moSize.Width + ", "
                + moSize.Height);

        // 
        // Polygon Inheritance Triangle
        //
        // XShape polyShape =
        // createShape(xDrawComponent,
        // new Point( 0, 0 ),
        // new Size( 0, 0 ),
        // "com.sun.star.drawing.ClosedBezierShape" );
        Object xObj = xDocumentFactory
                .createInstance("com.sun.star.drawing.ClosedBezierShape");
        XShape polyShape = (XShape) UnoRuntime.queryInterface(XShape.class,
                xObj);
        polyShape.setPosition(new Point(0, 0));
        polyShape.setSize(new Size(0, 0));

        // xDrawPage.add(polyShape);
        addShape(xObj, polyShape);
        // ucXShapes.add(polyShape);

        // Create the points
        PolyPolygonBezierCoords aCoords = new PolyPolygonBezierCoords();
        // allocating the outer sequence
        aCoords.Coordinates = new Point[1][];
        aCoords.Flags = new PolygonFlags[1][];

        int nPointCount = 3;
        Point[] pPolyPoints = new Point[nPointCount];
        PolygonFlags[] pPolyFlags = new PolygonFlags[nPointCount];

        for (int n = 0; n < nPointCount; n++) {
            pPolyPoints[n] = new Point();
        }

        int x = moPosition.X + (moSize.Width / 2) - 250;
        int y = moPosition.Y + moSize.Height + 500;

        pPolyPoints[0].X = x + 0;
        pPolyPoints[0].Y = y + 0;
        pPolyFlags[0] = PolygonFlags.NORMAL;
        pPolyPoints[1].X = x + 250;
        pPolyPoints[1].Y = y - 500;
        pPolyFlags[1] = PolygonFlags.NORMAL;
        pPolyPoints[2].X = x + 500;
        pPolyPoints[2].Y = y + 0;
        pPolyFlags[2] = PolygonFlags.NORMAL;
        // pPolyPoints[ 3 ].X = x + 0;
        // pPolyPoints[ 3 ].Y = y + 0;
        // pPolyFlags[ 3 ] = PolygonFlags.NORMAL;

        for (int i = 0; i < pPolyPoints.length; i++) {
            System.out.println("pPolyPoints[" + i + "]={" + pPolyPoints[i].X
                    + ", " + pPolyPoints[i].Y + "}");
        }

        //
        aCoords.Coordinates[0] = pPolyPoints;
        aCoords.Flags[0] = pPolyFlags;
        //

        setShapeProperty(xObj, "PolyPolygonBezier", aCoords);
        // White fill
        setShapeProperty(xObj, "FillColor", new Integer(0xFFFFFF));
        // move the shape to the back by changing the ZOrder
        // setShapeProperty(xObj, "ZOrder", new Integer( 1 ) );

        // Group UC and inheritance objects
        // XShapeGrouper ucgr =
        // (XShapeGrouper)UnoRuntime.queryInterface(XShapeGrouper.class,
        // xDoc);
        // XShape ucGroupXShape = ucgr.group(ucXShapes);

        return polyShape;
    }

    /**
     * Used to draw Inheritance Line
     */
    public XShape addRectLineXShape(XMultiServiceFactory xDocumentFactory,
            XShape from, XShape to) throws java.lang.Exception {
        //
        // Inhertiance Connector
        //
        Object connector = xDocumentFactory
                .createInstance("com.sun.star.drawing.ConnectorShape");
        com.sun.star.drawing.XShape xConnector = (com.sun.star.drawing.XShape) UnoRuntime
                .queryInterface(com.sun.star.drawing.XShape.class, connector);
        // xDrawPage.add(xConnector);
        addShape(connector, xConnector);

        setShapeProperty(connector, "StartShape", to);
        setShapeProperty(connector, "EndShape", from);

        // glue point positions: 0=top 1=left 2=bottom 3=right
        setShapeProperty(connector, "StartGluePointIndex", new Integer(2));
        setShapeProperty(connector, "EndGluePointIndex", new Integer(0));

        return xConnector;
    }

    public XShape addClassXShape(XMultiServiceFactory xDocumentFactory, int x,
            int y, int width, int height, String name)
            throws java.lang.Exception {

        // x = (int)(dx * x);
        // y = (int)(dy * y);
        // width = (int)(dx * width);
        // height = (int)(dy * height);

        //
        // Class Rectangle
        // 
        Object shape2 = xDocumentFactory
                .createInstance("com.sun.star.drawing.RectangleShape");
        XShape xShape2 = (XShape) UnoRuntime.queryInterface(XShape.class,
                shape2);
        xShape2.setPosition(new Point(x, y));
        xShape2.setSize(new Size(width, height));

        // xDrawPage.add(xShape2);
        addShape(shape2, xShape2);

        XPropertySet xShapePropSet2 = Soffice.getXPropertySet(xShape2);
        // setShapeProperty(shape2, "TextFitToSize",
        // TextFitToSizeType.PROPORTIONAL );
        setShapeProperty(shape2, "FillColor", new Integer(0xFFFFFF));
        // setShapeProperty(shape2, "TextLeftDistance", new Integer(500) );
        // setShapeProperty(shape2, "TextRightDistance", new Integer(500) );
        // setShapeProperty(shape2, "TextUpperDistance", new Integer(500) );
        // setShapeProperty(shape2, "TextLowerDistance", new Integer(500) );

        setShapeText(shape2, name);
        setShapeTextProperty(shape2, "ParaAdjust", ParagraphAdjust.CENTER);
        setShapeTextProperty(shape2, "CharColor", new Integer(
                SofficeDialog.COLOR_SUN1));
        setShapeTextProperty(shape2, "CharFontName", "Arial");
        setShapeTextProperty(shape2, "CharHeight", new Integer(10));

        return xShape2;
    }

    public XShape addUseCaseXShape(XMultiServiceFactory xDocumentFactory,
            int x, int y, int width, int height, String name)
            throws java.lang.Exception {

        // x = (int)(dx * x);
        // y = (int)(dy * y);
        // width = (int)(dx * width);
        // height = (int)(dy * height);

        //
        // Class Rectangle
        // 
        Object shape2 = xDocumentFactory
                .createInstance("com.sun.star.drawing.EllipseShape");
        XShape xShape2 = (XShape) UnoRuntime.queryInterface(XShape.class,
                shape2);
        xShape2.setPosition(new Point(x, y));
        xShape2.setSize(new Size(width, height));

        // xDrawPage.add(xShape2);
        addShape(shape2, xShape2);

        XPropertySet xShapePropSet2 = Soffice.getXPropertySet(xShape2);
        // setShapeProperty(shape2, "TextFitToSize",
        // TextFitToSizeType.PROPORTIONAL );
        setShapeProperty(shape2, "FillColor", new Integer(0xFFFFFF));
        // setShapeProperty(shape2, "TextLeftDistance", new Integer(500) );
        // setShapeProperty(shape2, "TextRightDistance", new Integer(500) );
        // setShapeProperty(shape2, "TextUpperDistance", new Integer(500) );
        // setShapeProperty(shape2, "TextLowerDistance", new Integer(500) );

        setShapeText(shape2, name);
        setShapeTextProperty(shape2, "ParaAdjust", ParagraphAdjust.CENTER);
        setShapeTextProperty(shape2, "CharColor", new Integer(
                SofficeDialog.COLOR_SUN1));
        setShapeTextProperty(shape2, "CharFontName", "Arial");
        setShapeTextProperty(shape2, "CharHeight", new Integer(10));

        return xShape2;
    }

}
