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

import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.Rectangle;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Enumeration;
import java.util.Hashtable;
import java.util.Iterator;
import java.util.List;
import java.util.Vector;

import javax.swing.BorderFactory;
import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JTextField;
import javax.swing.event.ChangeEvent;
import javax.swing.event.ChangeListener;

import org.apache.log4j.Logger;
import org.argouml.application.api.Argo;
import org.argouml.kernel.Project;
import org.argouml.kernel.ProjectManager;
import org.argouml.model.CoreHelper;
import org.argouml.model.Facade;
import org.argouml.model.Model;
import org.argouml.model.ModelManagementHelper;
import org.argouml.model.UseCasesHelper;
import org.argouml.ui.ProjectBrowser;
import org.argouml.uml.diagram.ArgoDiagram;
import org.argouml.uml.diagram.deployment.ui.FigComponent;
import org.argouml.uml.diagram.deployment.ui.UMLDeploymentDiagram;
import org.argouml.uml.diagram.ui.FigAssociation;
import org.argouml.uml.diagram.ui.FigDependency;
import org.argouml.uml.diagram.ui.FigEmptyRect;
import org.argouml.uml.diagram.ui.FigPermission;
import org.argouml.uml.diagram.ui.FigUsage;
import org.argouml.uml.diagram.use_case.ui.FigActor;
import org.argouml.uml.diagram.use_case.ui.FigInclude;
import org.argouml.uml.diagram.use_case.ui.FigUseCase;
import org.argouml.uml.diagram.use_case.ui.UMLUseCaseDiagram;
import org.argouml.util.ArgoDialog;
import org.tigris.gef.base.LayerPerspective;
import org.tigris.gef.graph.GraphEdgeRenderer;
import org.tigris.gef.graph.GraphModel;
import org.tigris.gef.presentation.AnnotationStrategy;
import org.tigris.gef.presentation.ArrowHead;
import org.tigris.gef.presentation.ArrowHeadDiamond;
import org.tigris.gef.presentation.ArrowHeadGreater;
import org.tigris.gef.presentation.ArrowHeadTriangle;
import org.tigris.gef.presentation.Fig;
import org.tigris.gef.presentation.FigCircle;
import org.tigris.gef.presentation.FigEdge;
import org.tigris.gef.presentation.FigGroup;
import org.tigris.gef.presentation.FigLine;
import org.tigris.gef.presentation.FigPoly;
import org.tigris.gef.presentation.FigRect;
import org.tigris.gef.presentation.FigText;

import com.sun.star.awt.Size;
import com.sun.star.beans.XPropertySet;
import com.sun.star.drawing.XShape;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.text.XText;
import com.sun.star.text.XTextCursor;
import com.sun.star.text.XTextDocument;
import com.sun.star.text.XTextTable;
import com.sun.star.uno.UnoRuntime;

/**
 * Dialog to generate StarOffice Document
 * 
 * @author <a href="mailto:matti@sun.com">Matti Pehrs</a>
 * @version $Id: SofficeDialog.java,v 1.7 2008/02/05 10:33:08 rastaman Exp $
 */
public class SofficeDialog extends ArgoDialog implements ActionListener {

    private Logger LOG = Logger.getLogger(SofficeDialog.class);

    JCheckBox drawingsCB = null;

    JCheckBox docCB = null;

    JCheckBox generateCurrentDiagramOnly = null;

    JCheckBox fixedScaleCB = null;

    JTextField scaleText = null;

    JTextField fontFamilyText = null;

    JCheckBox fixedFontSizeCB = null;

    JTextField fontSizeText = null;

    JCheckBox dictCB = null;

    JCheckBox dbgCB = null;

    JTextField outFileText = null;

    JButton fileBtn = null;

    // R G B
    // sun1 = 51 51 102 0x333366
    // sun2 = 102 102 153 0x666699
    // sun3 = 153 153 204 0x9999CC
    // sun4 = 204 204 255 0xCCCCFF
    public final static int COLOR_SUN1 = 0x333366;

    public final static int COLOR_SUN2 = 0x666699;

    public final static int COLOR_SUN3 = 0x9999CC;

    public final static int COLOR_SUN4 = 0xCCCCFF;

    public SofficeDialog() {
        super("Star/OpenOffice Plugin", ArgoDialog.OK_CANCEL_OPTION,
                true);
        init();
    }

    protected void nameButtons() {
        // Noting yet...
        JButton yBtn = getOkButton();
        yBtn.setText("Generate");

        JButton cBtn = getCancelButton();
        cBtn.setText("Cancel");
    }

    private void init() {

        JPanel panel = new JPanel();
        panel.setLayout(new BorderLayout());

        JPanel genCBPanel = new JPanel();
        genCBPanel.setBorder(BorderFactory.createTitledBorder(BorderFactory
                .createLineBorder(Color.black), "Generate"));
        GridBagLayout allGridbag = new GridBagLayout();
        genCBPanel.setLayout(allGridbag);

        JPanel drawCBPanel = new JPanel();
        drawCBPanel.setBorder(BorderFactory.createTitledBorder(BorderFactory
                .createLineBorder(Color.black), "Diagrams"));
        GridBagLayout drawGridbag = new GridBagLayout();
        drawCBPanel.setLayout(drawGridbag);

        JPanel docCBPanel = new JPanel();
        docCBPanel.setBorder(BorderFactory.createTitledBorder(BorderFactory
                .createLineBorder(Color.black), "Document"));
        GridBagLayout docGridbag = new GridBagLayout();
        docCBPanel.setLayout(docGridbag);

        JPanel miscCBPanel = new JPanel();
        miscCBPanel.setBorder(BorderFactory.createTitledBorder(BorderFactory
                .createLineBorder(Color.black), "Misc"));
        GridBagLayout miscGridbag = new GridBagLayout();
        miscCBPanel.setLayout(miscGridbag);

        GridBagConstraints gbc = new GridBagConstraints();

        int row = 1;

        drawingsCB = new JCheckBox("Generate Diagrams");
        drawingsCB.setSelected(true);
        drawingsCB.setEnabled(true);
        drawingsCB.addChangeListener(new ChangeListener() {
            public void stateChanged(ChangeEvent e) {
                boolean enable = drawingsCB.isSelected();
                generateCurrentDiagramOnly.setEnabled(enable);
                fixedScaleCB.setEnabled(enable);
                fontFamilyText.setEnabled(enable);
                fixedFontSizeCB.setEnabled(enable);
                if (enable) {
                    scaleText.setEnabled(fixedScaleCB.isSelected());
                    fontSizeText.setEnabled(fixedFontSizeCB.isSelected());
                } else {
                    scaleText.setEnabled(false);
                    fontSizeText.setEnabled(false);
                }
            }
        });
        gbc.fill = GridBagConstraints.BOTH;
        gbc.gridx = 0;
        gbc.gridy = row;
        gbc.gridwidth = 1;
        gbc.gridheight = 1;
        // gbc.weightx = 1.0;
        drawGridbag.setConstraints(drawingsCB, gbc);
        drawCBPanel.add(drawingsCB);

        generateCurrentDiagramOnly = new JCheckBox("Current Digram Only");
        generateCurrentDiagramOnly.setSelected(true);
        generateCurrentDiagramOnly.setEnabled(true);
        gbc.fill = GridBagConstraints.BOTH;
        gbc.gridx = 1;
        gbc.gridy = row;
        gbc.gridwidth = 1;
        gbc.gridheight = 1;
        // gbc.weightx = 1.0;
        drawGridbag.setConstraints(generateCurrentDiagramOnly, gbc);
        drawCBPanel.add(generateCurrentDiagramOnly);

        // -------------------------------------
        row++;

        docCB = new JCheckBox("Generate Document");
        docCB.setSelected(true);
        docCB.setEnabled(true);
        docCB.addChangeListener(new ChangeListener() {
            public void stateChanged(ChangeEvent e) {
                boolean enable = docCB.isSelected();
                fixedScaleCB.setEnabled(enable);
                fontFamilyText.setEnabled(enable);
                fixedFontSizeCB.setEnabled(enable);
                if (enable) {
                    scaleText.setEnabled(fixedScaleCB.isSelected());
                    fontSizeText.setEnabled(fixedFontSizeCB.isSelected());
                } else {
                    scaleText.setEnabled(false);
                    fontSizeText.setEnabled(false);
                }
            }
        });
        gbc.fill = GridBagConstraints.BOTH;
        gbc.gridx = 0;
        gbc.gridy = row;
        gbc.gridwidth = 1;
        gbc.gridheight = 1;
        // gbc.weightx = 1.0;
        docGridbag.setConstraints(docCB, gbc);
        docCBPanel.add(docCB);

        // -------------------------------------
        row++;

        fixedScaleCB = new JCheckBox("Fixed Scale");
        fixedScaleCB.setEnabled(true);
        fixedScaleCB.setSelected(false);
        fixedScaleCB.addChangeListener(new ChangeListener() {
            public void stateChanged(ChangeEvent e) {
                scaleText.setEnabled(fixedScaleCB.isSelected());
            }
        });
        gbc.fill = GridBagConstraints.BOTH;
        gbc.gridx = 0;
        gbc.gridy = row;
        gbc.gridwidth = 1;
        gbc.gridheight = 1;
        drawGridbag.setConstraints(fixedScaleCB, gbc);
        drawCBPanel.add(fixedScaleCB);

        scaleText = new JTextField(5);
        scaleText.setEnabled(false);
        scaleText.setText("30.0");
        gbc.fill = GridBagConstraints.BOTH;
        gbc.gridx = 1;
        gbc.gridy = row;
        gbc.gridwidth = 1;
        gbc.gridheight = 1;
        drawGridbag.setConstraints(scaleText, gbc);
        drawCBPanel.add(scaleText);

        // -------------------------------------
        row++;

        fixedFontSizeCB = new JCheckBox("Fixed Font Size");
        fixedFontSizeCB.setEnabled(true);
        fixedFontSizeCB.setSelected(false);
        fixedFontSizeCB.addChangeListener(new ChangeListener() {
            public void stateChanged(ChangeEvent e) {
                fontSizeText.setEnabled(fixedFontSizeCB.isSelected());
            }
        });
        gbc.fill = GridBagConstraints.BOTH;
        gbc.gridx = 0;
        gbc.gridy = row;
        gbc.gridwidth = 1;
        gbc.gridheight = 1;
        drawGridbag.setConstraints(fixedFontSizeCB, gbc);
        drawCBPanel.add(fixedFontSizeCB);

        fontSizeText = new JTextField(4);
        fontSizeText.setEnabled(false);
        fontSizeText.setText("8");
        gbc.fill = GridBagConstraints.BOTH;
        gbc.gridx = 1;
        gbc.gridy = row;
        gbc.gridwidth = 1;
        gbc.gridheight = 1;
        drawGridbag.setConstraints(fontSizeText, gbc);
        drawCBPanel.add(fontSizeText);

        // -------------------------------------
        row++;

        JLabel ffLabel = new JLabel("Font Family");
        ffLabel.setEnabled(true);
        gbc.fill = GridBagConstraints.BOTH;
        gbc.gridx = 0;
        gbc.gridy = row;
        gbc.gridwidth = 1;
        gbc.gridheight = 1;
        drawGridbag.setConstraints(ffLabel, gbc);
        drawCBPanel.add(ffLabel);

        fontFamilyText = new JTextField(4);
        fontFamilyText.setEnabled(true);
        fontFamilyText.setText("Dialog");
        gbc.fill = GridBagConstraints.BOTH;
        gbc.gridx = 1;
        gbc.gridy = row;
        gbc.gridwidth = 1;
        gbc.gridheight = 1;
        drawGridbag.setConstraints(fontFamilyText, gbc);
        drawCBPanel.add(fontFamilyText);

        // -------------------------------------
        row++;

        // Misc Panel

        row = 0;

        dictCB = new JCheckBox("Generate Dictionary");
        dictCB.setEnabled(true);
        dictCB.setSelected(false);
        gbc.fill = GridBagConstraints.BOTH;
        gbc.gridx = 0;
        gbc.gridy = row;
        gbc.gridwidth = 1;
        gbc.gridheight = 1;
        miscGridbag.setConstraints(dictCB, gbc);
        miscCBPanel.add(dictCB);

        dbgCB = new JCheckBox("Debug");
        dbgCB.setEnabled(true);
        dbgCB.setSelected(false);
        gbc.fill = GridBagConstraints.BOTH;
        gbc.gridx = 1;
        gbc.gridy = row;
        gbc.gridwidth = 1;
        gbc.gridheight = 1;
        miscGridbag.setConstraints(dbgCB, gbc);
        miscCBPanel.add(dbgCB);

        // -------------------------------------
        row++;

        // all Panel
        row = 0;

        gbc.fill = GridBagConstraints.BOTH;
        gbc.gridx = 0;
        gbc.gridy = row;
        gbc.gridwidth = 1;
        gbc.gridheight = 1;
        allGridbag.setConstraints(drawCBPanel, gbc);
        genCBPanel.add(drawCBPanel);
        // -------------------------------------
        row++;

        gbc.fill = GridBagConstraints.BOTH;
        gbc.gridx = 0;
        gbc.gridy = row;
        gbc.gridwidth = 1;
        gbc.gridheight = 1;
        allGridbag.setConstraints(docCBPanel, gbc);
        genCBPanel.add(docCBPanel);
        // -------------------------------------
        row++;

        gbc.fill = GridBagConstraints.BOTH;
        gbc.gridx = 0;
        gbc.gridy = row;
        gbc.gridwidth = 1;
        gbc.gridheight = 1;
        allGridbag.setConstraints(miscCBPanel, gbc);
        genCBPanel.add(miscCBPanel);
        // -------------------------------------
        row++;

        // Add all-panel
        panel.add(genCBPanel, BorderLayout.CENTER);

        setContent(panel);
    }

    private Object/* MInclude */includedAsBase(Collection ucIncludes,
            Object/* MUseCase */uc) {
        for (Iterator it = ucIncludes.iterator(); it.hasNext();) {
            Object/* MInclude */inc = /* (MInclude) */it.next();
            if (/* inc.getBase() */Model.getFacade().getBase(inc) == uc)
                return inc;
        }
        return null;
    }

    private Object/* MInclude */includedAsAddition(Collection ucIncludes, /* MUseCase */
    Object uc) {
        for (Iterator it = ucIncludes.iterator(); it.hasNext();) {
            Object/* MInclude */inc = it.next();
            if (Model.getFacade().getBase(inc) == uc)
                return inc;
        }
        return null;
    }

    public void actionPerformed(ActionEvent event) {
        if (event.getSource() == getCancelButton()) {
            hide();
            return;
        }

        if (event.getSource() == getOkButton()) {
            try {
                if (drawingsCB.isSelected()) {
                    generateFigDrawings();
                }
                if (dictCB.isSelected()) {
                    generateDictionary();
                }
                if (docCB.isSelected()) {
                    generateDoc();
                }
                showFeedback("Generated Document Successfully!");
                hide();
            } catch (Exception ex) {
                LOG.info(ex.toString().getClass().getName() + ":"
                        + ex.toString(), ex);
                ex.printStackTrace();
                showFeedback(ex.toString().getClass().getName() + ":"
                        + ex.toString());
            }
        }

    }

    public void showFeedback(String msg) {

        ProjectBrowser projectBrowser = ProjectBrowser.getInstance();
        ArgoDialog dialog = new ArgoDialog(
                "Star/OpenOffice Plugin", ArgoDialog.CLOSE_OPTION, true);

        dialog.setContent(new JLabel(msg));

        dialog.show();
    }

    public String toString(Collection coll, String id) {

        if (coll == null)
            return "";

        StringBuffer out = new StringBuffer();
        for (Iterator it = coll.iterator(); it.hasNext();) {
            java.lang.Object val = it.next();
            out.append("" + id + "[" + val.getClass().getName() + "]=" + val
                    + "\n");
        }
        return out.toString();
    }

    private void logCollection(Collection coll, String id) {
        if (coll == null)
            return;

        for (Iterator it = coll.iterator(); it.hasNext();) {
            java.lang.Object val = it.next();
            LOG.info("    " + id + "[" + val.getClass().getName() + "]=" + val);
        }
    }

    //
    // Staroffice
    //

    // Model Name
    String modelName = "Model: ?";

    // UseCase
    UseCasesHelper ucHelper = null;

    Collection useCases = null; // [MUseCase]

    // UseCase Includes
    Collection ucIncludes = null; // [MInclude]

    // Contains all UseCase Tags except "documentation"
    Vector useCasetags = null; // [String tag]

    // Componenets
    Collection components = null; // [MComponent]

    private void sofficeInit() {
        Object model = ProjectManager.getManager().getCurrentProject()
                .getModel();
        Facade facade = Model.getFacade();
        ucHelper = Model.getUseCasesHelper();
        useCases = ucHelper.getAllUseCases(model);
        if (useCases != null && useCases.size() > 0) {
            Object uc1 = useCases.iterator().next();
            modelName = "" + facade.getName(facade.getNamespace(uc1));
        }

        ModelManagementHelper mdlMgmt = Model.getModelManagementHelper();
        // Collection subSystems = mdlMgmt.getAllSubSystems();
        components = mdlMgmt.getAllModelElementsOfKind(model, Model
                .getMetaTypes().getComponent());
        ucIncludes = mdlMgmt.getAllModelElementsOfKind(model, Model
                .getMetaTypes().getInclude());

        // Trace UseCases for all tagged Values (except documentation)
        useCasetags = new Vector();
        for (Iterator it = useCases.iterator(); it.hasNext();) {
            Object uc = it.next();
            Collection tvals = facade.getTaggedValuesCollection(uc);
            for (Iterator tvit = tvals.iterator(); tvit.hasNext();) {
                Object tval = tvit.next();
                String tag = facade.getTag(tval);
                if (!"documentation".equals(tag) && !useCasetags.contains(tag))
                    useCasetags.add(tag);
            }
        }

    }

    private void generateDoc() throws Exception {

        sofficeInit();

        LOG.info("-------------------------------");
        LOG.info("begin: generateDoc()");
        LOG.info("-------------------------------");

        // ProjectBrowser projectBrowser = ProjectBrowser.getInstance();
        ProjectManager projMgr = ProjectManager.getManager();
        Project project = projMgr.getCurrentProject();
        Vector diagrams = project.getDiagrams();

        java.io.File templateFile = new java.io.File(Argo.getDirectory()
                + File.separator + "templates" + File.separator + "argouml.stw");
        StringBuffer sTmp = new StringBuffer("file:///");
        sTmp.append(templateFile.getCanonicalPath().replace('\\', '/'));
        String templateDoc = sTmp.toString();

        SWriterDoc doc = new SWriterDoc();
        doc.connect(templateDoc);

        doc.addParagraph("" + project.getName(), "Heading 1");
        doc.addParagraph("Document Generated "
                + (new java.util.Date()).toString(), null);

        for (Iterator it = diagrams.iterator(); it.hasNext();) {
            ArgoDiagram diagram = (ArgoDiagram) it.next();
            LOG.info("diagram=" + diagram.getClass().getName());

            // Ignore all but the Use Case Diagrams...
            if (!(diagram instanceof UMLUseCaseDiagram))
                continue;

            doc.addParagraph("" + diagram.getName(), "Heading 2");

            // Calc width and height of Diagram
            int diagram_width = 0;
            int diagram_height = 0;
            for (Enumeration en = diagram.elements(); en.hasMoreElements();) {
                Object element = en.nextElement();
                if (element instanceof Fig) {
                    Fig fig = (Fig) element;

                    int x = fig.getX();
                    int y = fig.getY();
                    int w = fig.getWidth();
                    int h = fig.getHeight();

                    int mw = x + w;
                    if (mw > diagram_width)
                        diagram_width = mw;

                    int mh = y + h;
                    if (mh > diagram_height)
                        diagram_height = mh;
                }

            }

            int DIAGRAM_HEIGHT = 13000;
            if (diagram_height > DIAGRAM_HEIGHT) {
                DIAGRAM_HEIGHT = diagram_height;
            }

            LOG.info("diagram_width=" + diagram_width);
            LOG.info("diagram_height=" + diagram_height);
            float delta = 1.0f;
            if (diagram_width > diagram_height) {
                delta = (float) doc.getFrameWidth()
                        / (float) (diagram_width + (diagram_width / 5));
            } else {
                delta = (float) DIAGRAM_HEIGHT
                        / (float) (diagram_height + (diagram_height / 5));
            }
            LOG.info("Before draw cache");
            DrawCache cache = new DrawCache();

            // Create bg rect
            Object bgrectobj = doc.getDocumentFactory().createInstance(
                    "com.sun.star.drawing.RectangleShape");
            XShape bgrect = (XShape) UnoRuntime.queryInterface(XShape.class,
                    bgrectobj);
            bgrect.setPosition(new com.sun.star.awt.Point(0, 0));
            if (diagram_height > 13000) {
                bgrect.setSize(new Size(doc.getFrameWidth(), diagram_height));
            } else {
                bgrect.setSize(new Size(doc.getFrameWidth(), DIAGRAM_HEIGHT));
            }
            LOG.info("Add shape");
            cache.addShape(bgrectobj, bgrect);
            cache.setShapeProperty(bgrectobj, "FillColor",
                    new Integer(0xFFFFFF));
            cache.setShapeProperty(bgrectobj, "TextContourFrame", new Boolean(
                    false));

            Hashtable diagramShapes = new Hashtable(); // [Fig src, XShape
            // docShape]

            for (Enumeration en = diagram.elements(); en.hasMoreElements();) {
                Object element = en.nextElement();
                LOG.info("  .element=(" + element.getClass().getName() + ")="
                        + element);
                if (element instanceof FigUseCase) {
                    FigUseCase fig = (FigUseCase) element;
                    debug(fig);
                    int x = (int) (delta * fig.getX());
                    int y = (int) (delta * fig.getY());
                    int w = (int) (delta * fig.getWidth());
                    int h = (int) (delta * fig.getHeight());
                    Object/* MUseCase */uc = /* (MUseCase) */fig.getOwner();

                    LOG.info("x=" + x + ", y=" + y + ", w=" + w + ", h=" + h);
                    XShape uc1 = cache.addUseCaseXShape(doc
                            .getDocumentFactory(), x, y, w, h, ""
                            + Model.getFacade().getName(uc));

                    diagramShapes.put(fig, uc1);

                } else if (element instanceof FigActor) {
                    FigActor fig = (FigActor) element;

                    int x = (int) (delta * fig.getX());
                    int y = (int) (delta * fig.getY());
                    int w = (int) (delta * fig.getWidth());
                    int h = (int) (delta * fig.getHeight());
                    Object/* MActor */actor = /* (MActor) */fig.getOwner();
                    LOG.info("x=" + x + ", y=" + y + ", w=" + w + ", h=" + h);
                    XShape xactor = cache.addActorXShape(doc
                            .getDocumentFactory(), x, y, w, h, ""
                            + Model.getFacade().getName(actor));

                    diagramShapes.put(fig, xactor);

                } else if (element instanceof FigAssociation) {
                    FigAssociation fig = (FigAssociation) element;
                    Fig dest = fig.getDestFigNode();
                    Fig src = fig.getSourceFigNode();
                    LOG.info("Association dest=(" + dest.getClass().getName()
                            + ")" + dest + ", src=(" + src.getClass().getName()
                            + ")" + src);
                    XShape xsrc = (XShape) diagramShapes.get(src);
                    XShape xdest = (XShape) diagramShapes.get(dest);
                    if (xsrc != null && xdest != null) {
                        // Add association
                        XShape ass1 = cache.addAssociationXShape(doc
                                .getDocumentFactory(), xsrc, xdest);
                    }
                }
            }
            // Add the drawing to the document
            cache.addCache2Text(doc.getDocumentFactory(), doc.getXText());
            doc.addParagraphBrake();

        }

        LOG.info("-------------------------------");
        LOG.info("end: generateDoc()");
        LOG.info("-------------------------------");
    }

    private void generateFigDrawings() throws Exception {

        LOG.info("-------------------------------");
        LOG.info("begin: generateFigDrawings()");
        LOG.info("-------------------------------");

        ProjectManager projMgr = ProjectManager.getManager();
        Project project = projMgr.getCurrentProject();

        java.io.File templateFile = new java.io.File(Argo.getDirectory()
                + File.separator + "templates" + File.separator + "argouml.std");
        StringBuffer sTmp = new StringBuffer("file:///");
        sTmp.append(templateFile.getCanonicalPath().replace('\\', '/'));
        String templateDoc = sTmp.toString();

        // Create and connect to Soffice
        SDrawDoc doc = new SDrawDoc();
        doc.connect(templateDoc);

        // Set Properties
        doc.setDrawTextPropotionalText(!fixedFontSizeCB.isSelected());
        doc.setDrawTextfontSize(Integer.parseInt(fontSizeText.getText()));
        doc.setDrawTextFontFamily(fontFamilyText.getText());

        if (generateCurrentDiagramOnly.isSelected()) {

            ArgoDiagram diagram = project.getActiveDiagram();
            generateFigDrawing(doc, diagram);

        } else {
            Vector diagrams = project.getDiagrams();
            for (Iterator it = diagrams.iterator(); it.hasNext();) {
                ArgoDiagram diagram = (ArgoDiagram) it.next();

                generateFigDrawing(doc, diagram);

                doc.addNewDrawPage();
            }

        }

        LOG.info("-------------------------------");
        LOG.info("end: generateFigDrawings()");
        LOG.info("-------------------------------");
    }

    private void generateFigDrawing(SDrawDoc doc, ArgoDiagram diagram)
            throws java.lang.Exception {

        //
        // Calc width and height of Diagram
        //
        int diagram_width = 0;
        int diagram_height = 0;
        LOG.info("Calc width and height of Diagram");
        for (Enumeration en = diagram.elements(); en.hasMoreElements();) {
            java.lang.Object element = en.nextElement();
            if (element instanceof Fig) {
                Fig fig = (Fig) element;

                // debug("element", fig);

                int x = fig.getX();
                int y = fig.getY();
                int w = fig.getWidth();
                int h = fig.getHeight();

                int mw = x + w;
                if (mw > diagram_width)
                    diagram_width = mw;

                int mh = y + h;
                if (mh > diagram_height)
                    diagram_height = mh;
            } else {
                LOG.info("non fig element=(" + debugElement(element)+")");
            }
        }
        LOG.info("diagram_width=" + diagram_width);
        LOG.info("diagram_height=" + diagram_height);
        if (fixedScaleCB.isSelected()) {
            float scale = Float.parseFloat(scaleText.getText());
            doc.setDX(scale);
            doc.setDY(scale);
        } else {
            if (diagram_width > diagram_height) {
                float delta = (float) doc.getDrawWidth()
                        / (float) (diagram_width + (diagram_width / 5));
                LOG.info("delta=" + delta);
                doc.setDX(delta);
                doc.setDY(delta);
            } else {
                float delta = (float) doc.getDrawHeight()
                        / (float) (diagram_height + (diagram_height / 5));
                LOG.info("delta=" + delta);
                doc.setDX(delta);
                doc.setDY(delta);
            }
        }

        // Draw Use Case Page
        doc.setDrawPageName(diagram.getName());

        Hashtable diagramShapes = new Hashtable(); // [Fig src, XShape docShape]

        for (Enumeration en = diagram.elements(); en.hasMoreElements();) {
            java.lang.Object element = en.nextElement();

            if (element instanceof Fig) {
                Fig fig = (Fig) element;

                drawFig(doc, fig);
            } else {
                LOG.info("non fig element=(" + debugElement(element)+")");
            }

        }

        GraphModel gModel = diagram.getGraphModel();
        LayerPerspective layer = diagram.getLayer();
        GraphEdgeRenderer renderer = layer.getGraphEdgeRenderer();

        // java.util.Vector edges = diagram.getEdges();
        // for(Iterator nit=edges.iterator();
        // nit.hasNext();) {
        // java.lang.Object element = nit.next();
        // FigEdge figEdge = renderer.getFigEdgeFor(gModel, layer, element);
        // if(figEdge != null) {
        // System.out.println("figEdge.getFig()=("+figEdge.getFig().getClass().getName()+")"+figEdge.getFig());
        // drawFig(doc, figEdge.getFig());
        // }
        // }

        Collection nodes = diagram.getNodes();
        for (Iterator nit = nodes.iterator(); nit.hasNext();) {
            java.lang.Object element = nit.next();
            if (element instanceof Fig) {
                Fig fig = (Fig) element;
                LOG.info("fig node=(" + debugElement(element) + ")");
                drawFig(doc, fig);
            } else {
                LOG.info("non-fig node=(" + debugElement(element) + ")");
            }
        }

    }

    private void generateDrawings() throws Exception {

        LOG.info("-------------------------------");
        LOG.info("begin: generateDrawings()");
        LOG.info("-------------------------------");

        // ProjectBrowser projectBrowser = ProjectBrowser.getInstance();
        ProjectManager projMgr = ProjectManager.getManager();
        Project project = projMgr.getCurrentProject();
        Vector diagrams = project.getDiagrams();

        java.io.File templateFile = new java.io.File(Argo.getDirectory()
                + File.separator + "templates" + File.separator + "argouml.std");
        StringBuffer sTmp = new StringBuffer("file:///");
        sTmp.append(templateFile.getCanonicalPath().replace('\\', '/'));
        String templateDoc = sTmp.toString();

        // Create and connect to Soffice
        SDrawDoc doc = new SDrawDoc();
        doc.connect(templateDoc);

        for (Iterator it = diagrams.iterator(); it.hasNext();) {
            ArgoDiagram diagram = (ArgoDiagram) it.next();

            //
            // Calc width and height of Diagram
            //
            int diagram_width = 0;
            int diagram_height = 0;
            LOG.info("Digaram Figs----------");
            for (Enumeration en = diagram.elements(); en.hasMoreElements();) {
                java.lang.Object element = en.nextElement();
                if (element instanceof Fig) {
                    Fig fig = (Fig) element;

                    debug(fig);

                    int x = fig.getX();
                    int y = fig.getY();
                    int w = fig.getWidth();
                    int h = fig.getHeight();

                    int mw = x + w;
                    if (mw > diagram_width)
                        diagram_width = mw;

                    int mh = y + h;
                    if (mh > diagram_height)
                        diagram_height = mh;
                }

            }
            LOG.info("Digaram Figs----------");
            LOG.info("diagram_width=" + diagram_width);
            LOG.info("diagram_height=" + diagram_height);
            if (diagram_width > diagram_height) {
                float delta = (float) doc.getDrawWidth()
                        / (float) (diagram_width + (diagram_width / 5));
                doc.setDX(delta);
                doc.setDY(delta);
            } else {
                float delta = (float) doc.getDrawHeight()
                        / (float) (diagram_height + (diagram_height / 5));
                doc.setDX(delta);
                doc.setDY(delta);
            }

            if (diagram instanceof UMLDeploymentDiagram) {

                // Draw Use Case Page
                doc.setDrawPageName(diagram.getName());

                Hashtable diagramShapes = new Hashtable(); // [Fig src, XShape
                // docShape]

                for (Enumeration en = diagram.elements(); en.hasMoreElements();) {
                    java.lang.Object element = en.nextElement();
                    LOG.info("  components.element=("
                            + element.getClass().getName() + ")=" + element);

                    if (element instanceof FigComponent) {
                        FigComponent fig = (FigComponent) element;
                        /* MComponent */Object comp = fig.getOwner();

                        int x = fig.getX();
                        int y = fig.getY();
                        int w = fig.getWidth();
                        int h = fig.getHeight();

                        // LOG.info("x="+x+", y="+y+", w="+w+", h="+h);
                        XShape xcomp = doc.drawComponent(x, y, w, h, ""
                                + Model.getFacade().getName(comp));

                        diagramShapes.put(fig, xcomp);

                    } else if (element instanceof FigDependency) {
                        FigDependency fig = (FigDependency) element;
                        /* MDependency */Object dep = fig.getOwner();

                        Fig dest = fig.getDestFigNode();
                        Fig src = fig.getSourceFigNode();
                        XShape xsrc = (XShape) diagramShapes.get(src);
                        XShape xdest = (XShape) diagramShapes.get(dest);
                        if (xsrc != null && xdest != null) {
                            XShape compAss = doc.drawAssociation(xsrc, xdest,
                                    false, true, null, null, null);
                            Soffice.setDashedLine(compAss);
                        }
                    }

                }
                doc.addNewDrawPage();

            } else if (diagram instanceof UMLUseCaseDiagram) {

                // Draw Use Case Page
                doc.setDrawPageName(diagram.getName());

                Hashtable diagramShapes = new Hashtable(); // [Fig src, XShape
                // docShape]

                for (Enumeration en = diagram.elements(); en.hasMoreElements();) {
                    java.lang.Object element = en.nextElement();
                    LOG.info("  uc.element=(" + element.getClass().getName()
                            + ")=" + element);
                    if (element instanceof FigUseCase) {
                        FigUseCase fig = (FigUseCase) element;
                        int x = fig.getX();
                        int y = fig.getY();
                        int w = fig.getWidth();
                        int h = fig.getHeight();
                        Object/* MUseCase */uc = fig.getOwner();

                        // LOG.info("x="+x+", y="+y+", w="+w+", h="+h);
                        XShape uc1 = doc.drawUseCase(x, y, w, h, ""
                                + Model.getFacade().getName(uc));

                        diagramShapes.put(fig, uc1);

                    } else if (element instanceof FigActor) {
                        FigActor fig = (FigActor) element;

                        int x = fig.getX();
                        int y = fig.getY();
                        int w = fig.getWidth();
                        int h = fig.getHeight();
                        Object/* MActor */actor = fig.getOwner();
                        // LOG.info("x="+x+", y="+y+", w="+w+", h="+h);
                        XShape xactor = doc.drawActor(x, y, w, h, ""
                                + Model.getFacade().getName(actor));

                        diagramShapes.put(fig, xactor);

                    } else if (element instanceof FigAssociation) {
                        FigAssociation fig = (FigAssociation) element;
                        Fig dest = fig.getDestFigNode();
                        Fig src = fig.getSourceFigNode();
                        // LOG.info("Association
                        // dest=("+dest.getClass().getName()+")"+dest+",
                        // src=("+src.getClass().getName()+")"+src);
                        XShape xsrc = (XShape) diagramShapes.get(src);
                        XShape xdest = (XShape) diagramShapes.get(dest);
                        if (xsrc != null && xdest != null) {

                            Object/* MAssociation */assoc = fig.getOwner();

                            LOG.info("assoc.name="
                                    + Model.getFacade().getName(assoc));
                            Object/* MAssociationEnd */srcAssoc = new ArrayList(
                                    Model.getFacade().getConnections(assoc))
                                    .get(0);
                            Object/* MAssociationEnd */destAssoc = new ArrayList(
                                    Model.getFacade().getConnections(assoc))
                                    .get(1);
                            // TODO: N-ary associations ?
                            LOG.info(" -->srcAssoc.name="
                                    + Model.getFacade().getName(srcAssoc));
                            LOG.info(" -->destAssoc.name="
                                    + Model.getFacade().getName(destAssoc));

                            Collection roles = Model.getFacade()
                                    .getAssociationRoles(assoc);
                            for (Iterator oit = roles.iterator(); oit.hasNext();) {
                                java.lang.Object obj = oit.next();
                                LOG.info("assoc.role=("
                                        + obj.getClass().getName() + ")" + obj);
                            }
                            Collection links = Model.getFacade()
                                    .getLinks(assoc);
                            for (Iterator oit = links.iterator(); oit.hasNext();) {
                                java.lang.Object obj = oit.next();
                                LOG.info("assoc.link=("
                                        + obj.getClass().getName() + ")" + obj);
                            }

                            // Add association
                            XShape ass1 = doc.drawAssociation(xsrc, xdest,
                                    false, false, Model.getFacade().getName(
                                            assoc), Model.getFacade().getName(
                                            srcAssoc), Model.getFacade()
                                            .getName(destAssoc));
                        }
                    }
                }
                doc.addNewDrawPage();
            }
        }

        LOG.info("-------------------------------");
        LOG.info("end: generateDrawings()");
        LOG.info("-------------------------------");
    }

    private void generateDictionary() throws Exception {

        LOG.info("Connec to StarOffice");
        XMultiServiceFactory xMSF = null;
        try {
            xMSF = Soffice.connect("localhost", 8100);
        } catch (Exception ex) {
            throw new Exception(
                    "Could not connect, make sure Star/OpenOffice is started and configured!",
                    ex);
        }

        // Open Writer document
        XTextDocument myDoc = null;

        // Load template
        java.io.File templateFile = new java.io.File(Argo.getDirectory()
                + File.separator + "templates" + File.separator + "argouml.stw");
        StringBuffer sTmp = new StringBuffer("file:///");
        sTmp.append(templateFile.getCanonicalPath().replace('\\', '/'));
        String templateDoc = sTmp.toString();

        LOG.info("Opening Writer Doc " + templateDoc);
        try {
            myDoc = Soffice.openWriter(xMSF, templateDoc);
        } catch (Exception ex) {
            throw new Exception("Could not open template "
                    + templateFile.getCanonicalPath() + "!", ex);
        }

        // getting the text object
        XText oText = myDoc.getText();
        // create a cursor object
        XTextCursor oCursor = oText.createTextCursor();

        //
        // Get Model stuff
        //
        Object model = ProjectManager.getManager().getCurrentProject()
                .getModel();
        ModelManagementHelper mdlMgmt = Model.getModelManagementHelper();
        Collection ns = mdlMgmt.getAllNamespaces(model);
        LOG.info("modelmgr.namespaces=" + ns);
        Collection subSystems = mdlMgmt.getAllSubSystems(model);
        Collection components = mdlMgmt.getAllModelElementsOfKind(model, Model
                .getMetaTypes().getComponent());
        Collection ucIncludes = mdlMgmt.getAllModelElementsOfKind(model, Model
                .getMetaTypes().getInclude());

        CoreHelper coreHelper = Model.getCoreHelper();
        Collection comps = coreHelper.getAllComponents(model);
        LOG.info("argo.components=" + comps);
        Collection classes = coreHelper.getAllClasses(model);
        LOG.info("argo.classes=" + classes);

        UseCasesHelper ucHelper = Model.getUseCasesHelper();
        Collection useCases = ucHelper.getAllUseCases(model);
        String modelName = "Model: ?";
        if (useCases != null && useCases.size() > 0) {
            Object/* MUseCase */uc1 = useCases.iterator().next();
            modelName = ""
                    + Model.getFacade().getName(
                            Model.getFacade().getNamespace(uc1));
        }
        // Trace UseCases for all tagged Values (except documentation)
        Vector tags = new Vector();
        for (Iterator it = useCases.iterator(); it.hasNext();) {
            Object/* MUseCase */uc = it.next();
            Collection tvals = Model.getFacade().getTaggedValuesCollection(uc);
            for (Iterator tvit = tvals.iterator(); tvit.hasNext();) {
                Object tval = /* (MTaggedValue) */tvit.next();
                String tag = Model.getFacade().getTag(tval);
                if (!"documentation".equals(tag) && !tags.contains(tag))
                    tags.add(tag);
            }
        }

        //
        // Heading: Model Name
        // 
        // -------------------------------------------------------------------------
        oText.insertString(oCursor, "" + modelName, false);
        Soffice.setLastParagraphStyle(oText, "Heading 1");
        Soffice.insertParagraphBreak(oText, oCursor);

        // Insert Generation date
        oText.insertString(oCursor, "Document Generated "
                + (new java.util.Date()).toString(), false);
        Soffice.insertParagraphBreak(oText, oCursor);

        //
        // USE CASES
        // 
        // insert a text table with the use cases.
        // -------------------------------------------------------------------------
        if (useCases != null && useCases.size() > 0) {
            oText.insertString(oCursor, "UseCases", false);
            Soffice.setLastParagraphStyle(oText, "Heading 2");
            Soffice.insertParagraphBreak(oText, oCursor);

            LOG.info("inserting a text table with Use Cases");

            // create instance of a text table
            XTextTable TT = Soffice.createXTextTable(myDoc);

            LOG.info("useCases.size()=" + useCases.size());
            LOG.info("tags.size()=" + tags.size() + ", tags=" + tags);
            // initialize the text table to the right size
            // int ucCount = 0;
            // for(Iterator cit = useCases.iterator();
            // cit.hasNext();) {
            // MUseCase uc = (MUseCase)cit.next();
            // ucCount += countNumUseCases(uc);
            // }

            TT.initialize(1 + useCases.size(), // rows
                    4 + tags.size()); // cols

            // insert the table into the text
            oText.insertTextContent(oCursor, TT, false);

            XPropertySet headerRow = Soffice.getXTextTableRowPropertySet(TT, 0);

            // get the property set of the text table
            XPropertySet oTPS = Soffice.getXTextTablePropertySet(TT);

            // Change the BackColor
            Soffice.setBackTransparent(headerRow, false);
            Soffice.setBackColor(headerRow, COLOR_SUN2);
            Soffice.setBackTransparent(oTPS, false);
            Soffice.setBackColor(oTPS, COLOR_SUN4);

            // write Text in the Table headers
            LOG.info("write text in the table headers");

            int headColor = Soffice.COLOR_WHITE;

            // Set Header Values
            Soffice.setCharColor(Soffice.insertText2Cell(TT, "A1", "Name"),
                    headColor);
            Soffice.setCharColor(Soffice.insertText2Cell(TT, "B1",
                    "Participants"), headColor);
            Soffice.setCharColor(Soffice.insertText2Cell(TT, "C1", "Includes"),
                    headColor);
            Soffice.setCharColor(Soffice.insertText2Cell(TT, "D1",
                    "Documentation"), headColor);
            int step = 0;
            for (Iterator it = tags.iterator(); it.hasNext(); step++) {
                String tag = (String) it.next();

                String cellId = ""
                        + (new Character((char) ('E' + (char) step))) + "1";
                Soffice.setCharColor(Soffice.insertText2Cell(TT, cellId, tag),
                        headColor);
            }

            // Go through all Use Cases
            // Start table on row 2
            int row = 2;
            for (Iterator it = useCases.iterator(); it.hasNext(); row++) {

                Object/* MUseCase */uc = it.next();

                // Ignore all included UCs, they will
                // be recursivly handled by uc2soffice()
                if (includedAsBase(ucIncludes, uc) != null)
                    continue;

                // Recursive interation on use-cases
                uc2soffice(TT, row, tags, uc, "");
            }
            // -------------------------------------------------------------------------

            //
            //
            // ACTORS
            // 
            // insert a text table with the Actors.
            // -------------------------------------------------------------------------

            oText.insertString(oCursor, "Actors", false);
            Soffice.setLastParagraphStyle(oText, "Heading 2");
            Soffice.insertParagraphBreak(oText, oCursor);

            LOG.info("inserting a text table with Actors");
            java.util.Collection actors = ucHelper.getAllActors(model);
            LOG.info("actors=" + actors);

            // Create table for actors
            TT = Soffice.createXTextTable(myDoc);

            // initialize the text table to the right size
            TT.initialize(1 + actors.size(), // rows
                    3); // cols

            // insert the table
            oText.insertTextContent(oCursor, TT, false);

            headerRow = Soffice.getXTextTableRowPropertySet(TT, 0);

            // get the property set of the text table
            oTPS = Soffice.getXTextTablePropertySet(TT);

            // Change the BackColor
            Soffice.setBackTransparent(headerRow, false);
            Soffice.setBackColor(headerRow, COLOR_SUN2);
            Soffice.setBackTransparent(oTPS, false);
            Soffice.setBackColor(oTPS, COLOR_SUN4);

            // write Text in the Table headers
            LOG.info("write text in the table headers");

            headColor = Soffice.COLOR_WHITE;

            // Set Header Values
            Soffice.setCharColor(Soffice.insertText2Cell(TT, "A1", "Name"),
                    headColor);
            Soffice.setCharColor(Soffice.insertText2Cell(TT, "B1",
                    "Participants"), headColor);
            Soffice.setCharColor(Soffice.insertText2Cell(TT, "C1",
                    "Documentation"), headColor);

            // Do the Actors
            int rowNum = 2;
            int rowColor = Soffice.COLOR_BLACK;
            for (Iterator it = actors.iterator(); it.hasNext(); rowNum++) {

                Object/* MActor */actor = it.next();

                String actorName = Model.getFacade().getName(actor);
                String actorParts = "";
                // TODO: is it ok in mdr facade ?
                Collection associationEnds = Model.getFacade()
                        .getAssociationEnds(actor);
                if (associationEnds != null) {
                    for (Iterator ait = associationEnds.iterator(); ait
                            .hasNext();) {
                        Object/* MAssociationEnd */end = ait.next();
                        Object/* MAssociationEnd */oend = Model.getFacade()
                                .getOppositeEnd(end);
                        if (actorParts.length() > 0)
                            actorParts += ",\n";
                        // TODO: is it ok for mdr ?
                        actorParts += Model.getFacade().getName(
                                Model.getFacade().getType(oend));
                    }
                }

                String actorDoc = Model.getFacade().getTaggedValueValue(actor,
                        "documentation");

                Soffice.setCharColor(Soffice.insertText2Cell(TT, "A" + rowNum,
                        actorName), rowColor);
                Soffice.setCharColor(Soffice.insertText2Cell(TT, "B" + rowNum,
                        actorParts), rowColor);
                Soffice.setCharColor(Soffice.insertText2Cell(TT, "C" + rowNum,
                        actorDoc), rowColor);
            }

            // -------------------------------------------------------------------------
        }

        //
        // Components
        //
        // -------------------------------------------------------------------------
        if (components != null && components.size() > 0) {

            LOG.info("Model Components.size = " + components.size());
            LOG.info("Model Components = " + components);

            oText.insertString(oCursor, "Components", false);
            Soffice.setLastParagraphStyle(oText, "Heading 2");
            Soffice.insertParagraphBreak(oText, oCursor);

            LOG.info("inserting a text table with Components");
            // create instance of a text table
            XTextTable TT = Soffice.createXTextTable(myDoc);

            LOG.info("initialize the text table to the right size");
            TT.initialize(1 + components.size(), // rows
                    2); // cols

            LOG.info("Insert the table");
            oText.insertTextContent(oCursor, TT, false);

            LOG.info("Get first rows propertyset");
            XPropertySet headerRow = Soffice.getXTextTableRowPropertySet(TT, 0);

            LOG.info("get the property set of the text table");
            XPropertySet oTPS = Soffice.getXTextTablePropertySet(TT);

            LOG.info("Change the BackColor");
            Soffice.setBackTransparent(headerRow, false);
            Soffice.setBackColor(headerRow, COLOR_SUN2);
            Soffice.setBackTransparent(oTPS, false);
            Soffice.setBackColor(oTPS, COLOR_SUN4);

            LOG.info("write text in the table headers");
            int headColor = Soffice.COLOR_WHITE;

            Soffice.setCharColor(Soffice.insertText2Cell(TT, "A1", "Name"),
                    headColor);
            Soffice.setCharColor(Soffice.insertText2Cell(TT, "B1",
                    "Documentation"), headColor);

            int row = 2;
            int rowColor = Soffice.COLOR_BLACK;
            for (Iterator it = components.iterator(); it.hasNext(); row++) {
                Object/* MComponent */comp = it.next();

                String compName = Model.getFacade().getName(comp);
                String compDoc = Model.getFacade().getTaggedValueValue(comp,
                        "documentation");

                Soffice.setCharColor(Soffice.insertText2Cell(TT, "A" + row,
                        compName), rowColor);
                Soffice.setCharColor(Soffice.insertText2Cell(TT, "B" + row,
                        compDoc), rowColor);
            }
        }
        // -------------------------------------------------------------------------

        // if the document should be disposed remove the slashes in front of the
        // next line
        // myDoc.dispose();

    }

    private void uc2soffice(XTextTable TT, int row, Collection tags,
            Object/* MUseCase */uc, String prefix) throws Exception {

        if (uc == null)
            return;

        String ucName = Model.getFacade().getName(uc);
        String ucActors = "";

        Collection associationEnds = Model.getFacade().getAssociationEnds(uc);
        if (associationEnds != null) {
            for (Iterator ait = associationEnds.iterator(); ait.hasNext();) {
                Object/* MAssociationEnd */end = ait.next();
                // LOG.info(" name="+end.getName());
                if (end != null) {
                    Object/* MAssociationEnd */oend = Model.getFacade()
                            .getOppositeEnd(end);
                    // LOG.info(" oend.type="+oend.getType());
                    if (ucActors.length() > 0)
                        ucActors += ", \n";
                    ucActors += Model.getFacade().getName(
                            Model.getFacade().getType(oend));
                }
            }
        }

        // Get Includes
        Collection includes = Model.getFacade().getIncludes(uc);
        String includesText = "";
        for (Iterator iit = includes.iterator(); iit.hasNext();) {
            Object/* MInclude */inc = iit.next();
            // row = uc2soffice(TT, row, tags, inc.getBase(), prefix+"->");
            if (Model.getFacade().getBase(inc) != null) {
                if (includesText.length() > 0)
                    includesText += ", \n";
                includesText += Model.getFacade().getBase(inc).toString();
            }
        }

        String ucDoc = Model.getFacade().getTaggedValueValue(uc,
                "documentation");

        int rowColor = Soffice.COLOR_BLACK;

        Soffice.setCharColor(Soffice.insertText2Cell(TT, "A" + row, prefix
                + ucName), rowColor);
        Soffice.setCharColor(Soffice.insertText2Cell(TT, "B" + row, ucActors),
                rowColor);
        Soffice.setCharColor(Soffice.insertText2Cell(TT, "C" + row,
                includesText), rowColor);
        Soffice.setCharColor(Soffice.insertText2Cell(TT, "D" + row, ucDoc),
                rowColor);
        int step = 0;
        for (Iterator ttit = tags.iterator(); ttit.hasNext(); step++) {
            String tag = (String) ttit.next();
            String val = Model.getFacade().getTaggedValueValue(uc, tag);
            String cellId = "" + (new Character((char) ('E' + (char) step)))
                    + "" + row;
            Soffice.setCharColor(Soffice.insertText2Cell(TT, cellId, val),
                    rowColor);
        }

    }

    /**
     * Recursively iterate through a use case (and its included use cases)
     */
    private int uc2sofficeRecursive(XTextTable TT, int row, Collection tags,
            Object/* MUseCase */uc, String prefix) throws Exception {

        if (uc == null)
            return row;

        String ucName = Model.getFacade().getName(uc);
        String ucActors = "";

        Collection associationEnds = Model.getFacade().getAssociationEnds(uc);
        if (associationEnds != null) {
            for (Iterator ait = associationEnds.iterator(); ait.hasNext();) {
                Object/* MAssociationEnd */end = ait.next();
                // LOG.info(" name="+end.getName());
                if (end != null) {
                    Object/* MAssociationEnd */oend = Model.getFacade()
                            .getOppositeEnd(end);
                    // LOG.info(" oend.type="+oend.getType());
                    if (ucActors.length() > 0)
                        ucActors += ", \n";
                    ucActors += Model.getFacade().getType(oend);
                }
            }
        }

        String ucDoc = Model.getFacade().getTaggedValueValue(uc,
                "documentation");

        int rowColor = Soffice.COLOR_BLACK;

        Soffice.setCharColor(Soffice.insertText2Cell(TT, "A" + row, prefix
                + ucName), rowColor);
        Soffice.setCharColor(Soffice.insertText2Cell(TT, "B" + row, ucActors),
                rowColor);
        Soffice.setCharColor(Soffice.insertText2Cell(TT, "C" + row, ucDoc),
                rowColor);
        int step = 0;
        for (Iterator ttit = tags.iterator(); ttit.hasNext(); step++) {
            String tag = (String) ttit.next();
            String val = Model.getFacade().getTaggedValueValue(uc, tag);
            String cellId = "" + (new Character((char) ('D' + (char) step)))
                    + "" + row;
            Soffice.setCharColor(Soffice.insertText2Cell(TT, cellId, val),
                    rowColor);
        }

        Collection includes = Model.getFacade().getIncludes(uc);
        LOG.info("Included uc(" + ucName + "):");
        for (Iterator iit = includes.iterator(); iit.hasNext();) {
            Object/* MInclude */inc = iit.next();
            LOG.info("include base=" + Model.getFacade().getBase(inc)
                    + ", addition=" + Model.getFacade().getAddition(inc));
            row = uc2sofficeRecursive(TT, row, tags, Model.getFacade().getBase(
                    inc), prefix + "->");
        }
        return row + 1;
    }

    /**
     * Recursively iterate through a use case (and its included use cases)
     */
    private int countNumUseCases(Object/* MUseCase */uc) throws Exception {

        if (uc == null)
            return 0;

        int count = 1;

        Collection includes = Model.getFacade().getIncludes(uc);
        for (Iterator iit = includes.iterator(); iit.hasNext();) {
            Object/* MInclude */inc = iit.next();
            count += countNumUseCases(Model.getFacade().getBase(inc));
        }
        return count;
    }

    private String toString(java.lang.Object list[]) {
        StringBuffer out = new StringBuffer();
        out.append("{");
        for (int i = 0; i < list.length; i++) {
            if (i > 0)
                out.append(", ");
            out.append(list[i].toString());
        }
        out.append("}");
        return out.toString();
    }

    private String toString(int list[]) {
        StringBuffer out = new StringBuffer();
        out.append("{");
        for (int i = 0; i < list.length; i++) {
            if (i > 0)
                out.append(", ");
            out.append("" + list[i]);
        }
        out.append("}");
        return out.toString();
    }

    private void debug(Fig fig) {
        debug("", fig);
    }

    private void drawFig(SDrawDoc doc, Fig fig) throws java.lang.Exception {

        if (!fig.isVisible())
            return;
        
        int x = fig.getX();
        int y = fig.getY();
        int w = fig.getWidth();
        int h = fig.getHeight();

        if (fig instanceof FigCircle) {
            LOG.info("Draw circle fig " + fig);
            FigCircle circle = (FigCircle) fig;
            XShape ellipse = doc.drawEllipse(x, y, w, h);
        } else if (fig instanceof FigText) {
            FigText text = (FigText) fig;

            LOG.info("Draw text=" + text.getText());


            String ftext = text.getText();

            XShape xtext = doc.drawText(x, y, w, h, ftext,
                    text.getFontFamily(), text.getFontSize());
        } else if (fig instanceof FigRect) {
            if (fig instanceof FigEmptyRect)
                return;
            FigRect rect = (FigRect) fig;
            Rectangle r = rect.getBounds();
            XShape xrect = doc.drawRect(r.x, r.y, r.width, r.height);
        } else if (fig instanceof FigPoly) {
            LOG.info("Draw poly");
            // FIXME
            FigPoly poly = (FigPoly) fig;
            List points = poly.getPointsList();
            doc.drawPoly(points);

        } else if (fig instanceof FigEdge) {
            LOG.info("Draw edge");
            FigEdge edge = (FigEdge) fig;

            LOG.info("(" + edge.getClass().getName() + ")[" + x + "," + y + ","
                    + w + "," + h + "]");
            int xs[] = edge.getXs();
            int ys[] = edge.getYs();
            int numPoints = edge.getNumPoints();
            LOG.info("xs[]=" + toString(xs));
            LOG.info("ys[]=" + toString(ys));

            boolean dashed = false;
            if (edge instanceof FigInclude || edge instanceof FigDependency
                    || edge instanceof FigPermission
                    || edge instanceof FigPermission
                    || edge instanceof FigUsage) {
                dashed = true;
            }

            for (int i = 0; i < (numPoints - 1); i++) {
                doc.drawLine(xs[i], ys[i], xs[i + 1], ys[i + 1], dashed);
            }

            ArrowHead destHead = edge.getDestArrowHead();
            if (destHead instanceof ArrowHeadGreater) {
                int i = edge.getNumPoints() - 2;
                if (edge.getNumPoints() > 0) {
                    doc
                            .drawGreaterArrowHead(xs[i], ys[i], xs[i + 1],
                                    ys[i + 1]);
                } else {
                    LOG
                            .info("Somethings wrong destHead = ArrowHeadGreater but edge.getNumPoints()="
                                    + edge.getNumPoints());
                }
            } else if (destHead instanceof ArrowHeadTriangle) {
                int i = edge.getNumPoints() - 2;
                if (edge.getNumPoints() > 0) {
                    doc.drawTriangleArrowHead(xs[i], ys[i], xs[i + 1],
                            ys[i + 1]);
                } else {
                    LOG
                            .info("Somethings wrong destHead = ArrowHeadTriangle but edge.getNumPoints()="
                                    + edge.getNumPoints());
                }
            } else if (destHead instanceof ArrowHeadDiamond) {

                Color color = destHead.getFillColor();
                int red = color.getRed();
                int green = color.getGreen();
                int blue = color.getGreen();
                int fillColor = (0x010000 * red) + (0x000100 * green)
                        + (0x000001 * blue);
                int i = edge.getNumPoints() - 2;
                if (edge.getNumPoints() > 0) {
                    doc.drawDiamondArrowHead(xs[i], ys[i], xs[i + 1],
                            ys[i + 1], fillColor);
                } else {
                    LOG
                            .info("Somethings wrong destHead = ArrowHeadDiamond but edge.getNumPoints()="
                                    + edge.getNumPoints());
                }
            }

            ArrowHead srcHead = edge.getSourceArrowHead();
            if (srcHead instanceof ArrowHeadGreater) {
                doc.drawGreaterArrowHead(xs[1], ys[1], xs[0], ys[0]);
            } else if (srcHead instanceof ArrowHeadTriangle) {
                doc.drawTriangleArrowHead(xs[1], ys[1], xs[0], ys[0]);
            } else if (srcHead instanceof ArrowHeadDiamond) {
                Color color = srcHead.getFillColor();
                int red = color.getRed();
                int green = color.getGreen();
                int blue = color.getGreen();
                int fillColor = (0x010000 * red) + (0x000100 * green)
                        + (0x000001 * blue);
                doc.drawDiamondArrowHead(xs[1], ys[1], xs[0], ys[0], fillColor);
            }

            java.util.Vector pathFigs = edge.getPathItemFigs();
            for (Iterator it = pathFigs.iterator(); it.hasNext();) {
                Fig f = (Fig) it.next();
                drawFig(doc, f);
            }

        } else if (fig instanceof FigLine) {
            FigLine line = (FigLine) fig;
            doc.drawLine(line.getX1(), line.getY1(), line.getX2(), line.getY2());
        } else {
            LOG.info("No single presentation for " + fig);
        }

        if (fig instanceof FigGroup) {
            FigGroup fGroup = (FigGroup) fig;
            // Get contained figs
            LOG.info(">>>Draw contained figs of " + fig);

            List figs = fGroup.getFigs();
            int i = 0;
            for (Iterator it = figs.iterator(); it.hasNext(); i++) {
                Fig ffig = (Fig) it.next();
                drawFig(doc, ffig);
            }
            LOG.info("<<<End draw contained figs of " + fig);

        }

        Fig owner = fig.getAnnotationOwner();
        if (owner != null) {
            LOG.info("getAnnotationOwner=(" + owner.getClass().getName() + ")"
                    + owner);
        }

        AnnotationStrategy annStrategy = fig.getAnnotationStrategy();
        for (Enumeration en = annStrategy.getAllAnnotations(); en
                .hasMoreElements();) {
            java.lang.Object obj = en.nextElement();
            LOG.info("annotation obj=(" + obj.getClass().getName() + ")" + obj);
            if (obj instanceof Fig) {
                drawFig(doc, (Fig) obj);
            }
        }
    }

    private void debug(String pfix, Fig fig) {

        LOG.info(pfix + "-------------------------------------------------");
        LOG.info(pfix + ".class=" + fig.getClass().getName());
        LOG.info(pfix + ".id=" + fig.getId());
        // java.awt.Point getLocation();
        LOG.info(pfix + ".location=" + fig.getLocation());
        // LOG.info(pfix+".x="+fig.getX());
        // LOG.info(pfix+".y="+fig.getY());
        // java.awt.Dimension getSize();
        LOG.info(pfix + ".size=" + fig.getSize());
        // LOG.info(pfix+".width="+fig.getWidth());
        // LOG.info(pfix+".height="+fig.getHeight());

        // java.awt.Point getFirstPoint();
        LOG.info(pfix + ".firstPoint=" + fig.getFirstPoint());
        // java.awt.Point getLastPoint();
        LOG.info(pfix + ".lastPoint=" + fig.getLastPoint());
        // java.awt.Point[] getPoints();
        LOG.info(pfix + ".numPoints=" + fig.getNumPoints());
        LOG.info(pfix + ".points=" + toString(fig.getPoints()));

        // LOG.info(pfix+".halfHeight="+fig.getHalfHeight());
        // LOG.info(pfix+".halfWidth="+fig.getHalfWidth());
        // LOG.info(pfix+".context="+fig.getContext());
        // LOG.info(pfix+".dashed="+fig.getDashed());
        // LOG.info(pfix+".dashed01="+fig.getDashed01());
        // LOG.info(pfix+".dashedString="+fig.getDashedString());
        Vector enclosed = fig.getEnclosedFigs();
        if (enclosed == null) {
            LOG.info(pfix + ".enclosedFigs=null");
        } else {
            int i = 0;
            for (Iterator it = enclosed.iterator(); it.hasNext(); i++) {
                Fig eFig = (Fig) it.next();
                debug(pfix + ".enclosedFigs[" + i + "]=", eFig);
            }
        }
        // LOG.info(pfix+".fillColor="+fig.getFillColor());
        // LOG.info(pfix+".filled="+fig.getFilled());
        // LOG.info(pfix+".filled01="+fig.getFilled01());
        // Vector getGravityPoints();
        // Layer getLayer();
        LOG.info(pfix + ".layer=" + fig.getLayer());
        // java.awt.Color getLineColor();
        // LOG.info(pfix+".lineColor="+fig.getLineColor());
        // LOG.info(pfix+".lineWidth="+fig.getLineWidth());
        // LOG.info(pfix+".locked="+fig.getLocked());
        // java.lang.Object getOwner();
        // LOG.info(pfix+".owner="+fig.getOwner());
        // LOG.info(pfix+".perimeterLength="+fig.getPerimeterLength());
        // LOG.info(pfix+".privateData="+fig.getPrivateData());
        // LOG.info(pfix+".resource="+fig.getResource());
        // java.awt.Rectangle getTrapRect();
        // LOG.info(pfix+".trapRect="+fig.getTrapRect());
        // LOG.info(pfix+".useTrapRect="+fig.getUseTrapRect());
        // LOG.info(pfix+".visState="+fig.getVisState());
        // int[] getXs();
        // LOG.info(pfix+".xs="+fig.getXs());
        // int[] getYs();
        // LOG.info(pfix+".ys="+fig.getYs());
        // LOG.info(pfix+".isAnnotation="+fig.isAnnotation());
        // LOG.info(pfix+".isDisplayed="+fig.isDisplayed());
        // LOG.info(pfix+".isLowerRightResizable="+fig.isLowerRightResizable());
        // LOG.info(pfix+".isMovable="+fig.isMovable());
        // LOG.info(pfix+".isReshapable="+fig.isReshapable());
        // LOG.info(pfix+".isResizable="+fig.isResizable());
        // LOG.info(pfix+".isRotatable="+fig.isRotatable());

        // Fig getGroup()

        if (fig instanceof FigLine) {
            FigLine line = (FigLine) fig;
            LOG.info(pfix + ".x1=" + line.getX1());
            LOG.info(pfix + ".y1=" + line.getY1());
            LOG.info(pfix + ".x2=" + line.getX2());
            LOG.info(pfix + ".y2=" + line.getY2());
        }

        if (fig instanceof FigText) {
            FigText text = (FigText) fig;
            LOG.info(pfix + ".text=" + text.getText());
            LOG.info(pfix + ".fontFamily=" + text.getFontFamily());
            LOG.info(pfix + ".fontSize=" + text.getFontSize());
            LOG.info(pfix + ".bold=" + text.getBold());
            LOG.info(pfix + ".italic=" + text.getItalic());
            LOG.info(pfix + ".underline=" + text.getUnderline());
        }

        if (fig instanceof FigGroup) {
            FigGroup fGroup = (FigGroup) fig;
            // Get contained figs

            List figs = fGroup.getFigs();
            int i = 0;
            for (Iterator it = figs.iterator(); it.hasNext(); i++) {
                Fig ffig = (Fig) it.next();
                debug(pfix + ".fig[" + i + "]=", ffig);
            }
        }
        LOG.info(pfix + "-------------------------------------------------");
    }

    private String debugElement(Object element) {
        String s = null;
        if (Model.getFacade().isAModelElement(element)) {
            s = Model.getFacade().getUMLClassName(element)
                + "<<" + Model.getFacade().getName(element)+">>";
        } else {
            s = element.getClass().getName()+"<<"+element+">>";
        }
        return s;
    }
}
