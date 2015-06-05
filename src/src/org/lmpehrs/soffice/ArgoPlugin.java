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

import java.awt.Component;
import java.awt.event.ActionEvent;
import java.util.Vector;

import javax.swing.AbstractAction;
import javax.swing.JMenu;
import javax.swing.JMenuBar;
import javax.swing.JMenuItem;

import org.apache.log4j.Logger;
import org.argouml.moduleloader.ModuleInterface;
import org.argouml.ui.ArgoToolbarManager;
import org.argouml.ui.ProjectBrowser;
import org.argouml.ui.cmd.GenericArgoMenuBar;

/**
 * Plugin that exposes itself to all JMenuItem contexts and provides a
 * StarOffice Document Generation
 * 
 * @author <a href="mailto:matti@sun.com">Matti Pehrs</a>
 * @version $Id: ArgoPlugin.java,v 1.5 2008/02/05 10:33:08 rastaman Exp $
 */
public class ArgoPlugin extends AbstractAction implements ModuleInterface {

    private Logger LOG = Logger.getLogger(ArgoPlugin.class);

    /**
     * This is not publicly creatable.
     */
    public ArgoPlugin() {
        //super("Generate HTML", false);
    	LOG.info("StarOffice plugin instantiated");
    }

    // //////////////////////////////////////////////////////////////
    // Main methods.

    SofficeDialog dialog = null;

    /**
     * Just let the tester know that we got executed.
     */
    public void actionPerformed(ActionEvent event) {
        LOG.error("StarOffice Generator started...");
        // Argo.log.info("User clicked on '" + event.getActionCommand() + "'");

        if (dialog == null) {
            dialog = new SofficeDialog();
        }
        dialog.show();
        
    }

    public void setModuleEnabled(boolean v) {
    }

    public boolean initializeModule() {
        LOG.info("+--------------------------------------+");
        LOG.info("| StarOffice Plugin Generator enabled! |");
        LOG.info("+--------------------------------------+");
        
        return true;
    }

    public boolean enable() {
    	LOG.error("Enabling StarOffice module");    	
    	GenericArgoMenuBar.registerMenuItem(getMenu());
    	LOG.error("Enabled StarOffice module");
        return true;
    }

    public boolean disable() {
        return true;
    }

    public String getName() {
        return "SofficePlugin";
    }

    public String getModuleDescription() {
        return "StarOffice Documentation Generator";
    }

    public String getModuleAuthor() {
        return "Matti Pehrs";
    }

    public String getModuleVersion() {
        return "0.2";
    }

    public String getModuleWebsite() {
        return "http://argosoffice.tigris.org/";
    }
    
    public String getModuleKey() {
        return "module.language.generator.soffice";
    }

    public JMenu getMenu() {
    	JMenu menu = new JMenu("Export");
        JMenuItem _menuItem = new JMenuItem("Generate StarOffice...");
        _menuItem.addActionListener(this);
        menu.add(_menuItem);
        return menu;
    }
    
    /**
     *
     */
    public String getInfo(int i) {
    	switch (i) {
    	case ModuleInterface.DOWNLOADSITE:
    		return getModuleWebsite();
		case ModuleInterface.VERSION:
			return getModuleVersion();
		case ModuleInterface.AUTHOR:
			return getModuleAuthor();
		case ModuleInterface.DESCRIPTION:			
			return getModuleDescription();
		default:
			return getModuleDescription()+ " - "+getModuleVersion();
    	}
    }
}
