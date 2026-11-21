/*
 * Copyright 2011-2016 Jonathan Nash
 * Modified 2026 by infraLeap (infraleap.com): ported to Vaadin Flow.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package com.vaadin.addon.tableexport;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.Serializable;
import java.util.List;
import java.util.logging.Logger;

import com.vaadin.flow.component.UI;
import com.vaadin.flow.server.StreamRegistration;

public abstract class TableExport implements Serializable {

    private static final long serialVersionUID = -2972527330991334117L;
    private static final Logger LOGGER = Logger.getLogger(TableExport.class.getName());

    public static String EXCEL_MIME_TYPE = "application/vnd.ms-excel";
    public static String CSV_MIME_TYPE = "text/csv";

    /** The Tableholder to export. */
    private TableHolder tableHolder;

    /** The window to send the export result */
    protected String exportWindow = "_self";

    protected String mimeType;

    public TableExport(TableHolder tableHolder) {
        this.tableHolder = tableHolder;
    }

    public TableHolder getTableHolder() {
        return tableHolder;
    }

    public List<Object> getPropIds() {
        return tableHolder.getPropIds();
    }

    public void setTableHolder(final TableHolder tableHolder) {
        this.tableHolder = tableHolder;
    }

    public boolean isHierarchical() {
        return tableHolder.isHierarchical();
    }

    public abstract void convertTable();
    public abstract boolean sendConverted();

    /**
     * Create and export the Table contents as some sort of file type. In the case of conversion to
     * Excel it would be an ".xls" file containing the contents as a report. Only the export()
     * method needs to be called. If the user wishes to manipulate the converted object to export,
     * then convertTable() should be called separately, and, after manipulation, sendConverted().
     */

    public void export() {
        convertTable();
        sendConverted();
    }

    /**
     * Utility method to send the converted object to the user, if it has been written to a
     * temporary File.
     * 
     * Code obtained from: http://vaadin.com/forum/-/message_boards/view_message/159583
     * 
     * @return true, if successful
     */
    public boolean sendConvertedFileToUser(final UI app, final File fileToExport,
            final String exportFileName, final String mimeType) {
        setMimeType(mimeType);
        return sendConvertedFileToUser(app, fileToExport, exportFileName);

    }

    protected boolean sendConvertedFileToUser(final UI app, final File fileToExport,
            final String exportFileName) {
        TemporaryFileDownloadResource resource;
        try {
            resource = new TemporaryFileDownloadResource(exportFileName, mimeType, fileToExport);
            final UI ui = null == app ? UI.getCurrent() : app;
            final StreamRegistration registration =
                    ui.getSession().getResourceRegistry().registerResource(resource, ui.getElement());
            resource.setRegistration(registration);
            // the file comes as an attachment, so the browser stays on the page
            ui.getPage().open(registration.getResourceUri().toString(), exportWindow);
        } catch (final FileNotFoundException e) {
            LOGGER.warning("Sending file to user failed with FileNotFoundException " + e);
            return false;
        }
        return true;
    }

    public String getExportWindow() {
        return this.exportWindow;
    }

    public void setExportWindow(final String exportWindow) {
        this.exportWindow = exportWindow;
    }

    public String getMimeType() {
        return this.mimeType;
    }

    public void setMimeType(final String mimeType) {
        this.mimeType = mimeType;
    }

}
